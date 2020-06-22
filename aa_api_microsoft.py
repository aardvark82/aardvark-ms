import datetime
import time

import aa_api_data
import aa_flask_cache
import aa_globals
import aa_helper_methods
import aa_loggers
import aa_sqlalchemy
from aa_auth_microsoft import ms_settings, AuthMicrosoft
from requests_oauthlib import OAuth2Session
########### Settings import for different environments ##########
from settings.settings import *

if isProduction():
    from settings.settings_prod import *
elif isStaging():
    from settings.settings_staging import *
elif isStagingNew():
    from settings.settings_staging_static import *
else:
    isDevelopment()
    from settings.settings_dev import *

log = aa_loggers.logging.getLogger(__name__)

cache = aa_flask_cache.getCacheObject()


###################################################################

class ApiMicrosoft(aa_api_data.ApiData):
    def __init__(self):
        self.ms_auth_client = AuthMicrosoft()

    def APIgetDataSourceType(self):
        return 'microsoft'

    def APIgetLabelNamesForLabelIds(self, label_ids):
        return ['Microsoft']

    def graph_generator(self, graph_client, endpoint=None, params=None):
        while endpoint:
            log.info('Getting 10 emails from Microsoft Graph API')
            ms_api_response = graph_client.get(url=endpoint, params=params, headers={"Prefer": 'IdType="ImmutableId"'})
            if ms_api_response.status_code == 429:
                retry_after = ms_api_response.headers.get('Retry-After')
                log.warning(f"Microsoft Graph API limit reached. Sleep for {retry_after}")
                time.sleep(retry_after)
                continue
            elif ms_api_response.status_code == 200:
                self.messages_count = ms_api_response.json().get('@odata.count')
                yield from ms_api_response.json().get('value')
                endpoint = ms_api_response.json().get('@odata.nextLink')
                params = None

            else:
                log.warning(
                    f'Microsoft API error: Status code {ms_api_response.status_code}, error:{ms_api_response.json()}')
                return None

    def APIgetAllEmails(self, formatted_labels: str = None, date_first: datetime.datetime = None,
                        date_last: datetime.datetime = None,
                        filter: str = None, metadata_limit: int = None, user: str = None, output: bool = False):

        """
        Get all the messages for the default period (1 week)
        :return: if output, return list of dicts
        """
        self.messages = []

        if not date_first:
            date_first = aa_globals.todayLocalUser() - datetime.timedelta(days=7)

        elif not isinstance(date_first, datetime.date):
            date_first = aa_helper_methods.build_tz_aware_datetime(date_first, to_utc=False).date()

        if not date_last:
            date_last = aa_globals.todayLocalUser()

        elif not isinstance(date_last, datetime.date):
            date_last = aa_helper_methods.build_tz_aware_datetime(date_last, to_utc=False).date()

        date_last += datetime.timedelta(days=1)  # include the current day

        ms_token = self.ms_auth_client.get_or_refresh_token()
        graph_client = OAuth2Session(token=ms_token)

        endpoint = '{}/me/messages'.format(ms_settings.get('graph_url'))

        # Get emails
        query_params = {
            '$count': 'true',
            '$orderby': 'receivedDateTime DESC',
            '$filter': 'receivedDateTime ge {} AND receivedDateTime le {}'.format(date_first, date_last)
        }

        self.messages = self.graph_generator(graph_client=graph_client, endpoint=endpoint, params=query_params)

        if output:
            return self.parseAllEmailAPIResponses()

    def parseAllEmailAPIResponses(self, output=True):
        parsed_messages = []

        try:
            messages = self.messages

        except AttributeError:
            log.info('No messages to parse')

        else:
            count = 0
            for message in messages:
                count += 1
                parsed_message = None
                try:
                    parsed_message = aa_api_data.parse_message_from_microsoft(message)

                    self.processInsightsAndStoreInDbIfNeeded(
                        message_internal=parsed_message)

                    parsed_messages.append(parsed_message)

                except Exception as e:
                    error_message = "Exception - parseEmailResultsRequest  Exception code:{} Subject:{}"
                    if parsed_message:
                        log.warning(error_message.format(e.__str__(), parsed_message.get('Subject', 'unknown')))
                    else:
                        log.warning(error_message.format(e.__str__(), 'unknown'))

                # User feedback
                if count % 50 == 0:
                    log.info(
                        f"API {self.APIgetDataSourceType().title()} Parsing response {count} of {self.messages_count} "
                        f"messages")

            aa_sqlalchemy.proxyCommit()

        finally:
            if output:
                return parsed_messages
