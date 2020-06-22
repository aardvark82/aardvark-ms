import json
import os

import aa_credentials
import aa_database_queries
import aa_globals
import aa_loggers
import aa_users
import time
import yaml
from requests_oauthlib import OAuth2Session

log = aa_loggers.logging.getLogger(__name__)

os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
os.environ['OAUTHLIB_IGNORE_SCOPE_CHANGE'] = '1'

# Load config from YAML file
stream = open('microsoft_oauth_settings.yaml', 'r')
ms_settings = yaml.load(stream, yaml.SafeLoader)
authorize_url = '{0}{1}'.format(ms_settings['authority'], ms_settings['authorize_endpoint'])
token_url = '{0}{1}'.format(ms_settings['authority'], ms_settings['token_endpoint'])


class AuthMicrosoft():
    def __init__(self, scopes=None, redirect_uri: str = None):
        if scopes:
            aa_globals.session['auth_scopes_requested'] = scopes

        if redirect_uri:
            self.microsoft_redirect_uri = redirect_uri
        else:
            self.microsoft_redirect_uri = aa_globals.get_auth_callback_url(provider='microsoft')

        # get_sign_in_url
        ms_auth = OAuth2Session(ms_settings['app_id'],
                                scope=aa_globals.session.get('auth_scopes_requested'),
                                redirect_uri=self.microsoft_redirect_uri)

        self.sign_in_url, self.state = ms_auth.authorization_url(authorize_url, prompt='select_account')

    def AuthGetUserEmail(self):
        user = self.get_user()
        return user.get('mail') or user.get('userPrincipalName')

    def get_token_from_code(self, callback_url, expected_state):
        # Initialize the OAuth client
        ms_auth = OAuth2Session(ms_settings['app_id'],
                                state=expected_state,
                                scope=aa_globals.session.get('auth_scopes_requested'),
                                redirect_uri=self.microsoft_redirect_uri)

        token = ms_auth.fetch_token(token_url,
                                    client_secret=ms_settings['app_secret'],
                                    authorization_response=callback_url)

        return token

    def get_or_refresh_token(self):
        token = aa_globals.session.get('microsoft_auth_token')

        if not token:
            token = aa_credentials.get_provider_token_from_db(provider='microsoft')

        if token:
            now = time.time()
            expire_time = token['expires_at'] - 300

            if now < expire_time:
                return token

            elif token.get('refresh_token'):
                ms_auth = OAuth2Session(ms_settings['app_id'],
                                        token=token,
                                        scope=aa_globals.session.get('auth_scopes_requested'),
                                        redirect_uri=self.microsoft_redirect_uri
                                        )
                refresh_params = {
                    'client_id': ms_settings['app_id'],
                    'client_secret': ms_settings['app_secret'],
                }
                new_token = ms_auth.refresh_token(token_url, **refresh_params)

                self.store_token(new_token)

                return new_token

            else:  # if scopes doesn't allow to get a refresh_token
                return None

        else:
            return None

    def sign_in(self):
        # Save the expected state so we can validate in the callback
        aa_globals.session['microsoft_auth_state'] = self.state

        # Redirect to the Azure sign-in page
        return self.sign_in_url

    def request_consent(self, prompt):
        # get_sign_in_url
        ms_auth = OAuth2Session(ms_settings['app_id'],
                                scope=aa_globals.session.get('auth_scopes_requested'),
                                redirect_uri=self.microsoft_redirect_uri)
        self.sign_in_url, self.state = ms_auth.authorization_url(authorize_url, prompt=prompt)

        # Save the expected state so we can validate in the callback
        aa_globals.session['microsoft_auth_state'] = self.state

        # Redirect to the Azure sign-in page
        return self.sign_in_url

    def get_user(self):
        token = self.get_or_refresh_token()
        graph_client = OAuth2Session(token=token)
        user = graph_client.get(ms_settings['graph_url'] + '/me')
        return user.json()

    def AuthGetCurrentProfilePictureForUser(self):
        import base64
        token = self.get_or_refresh_token()
        graph_client = OAuth2Session(token=token)
        response = graph_client.get(ms_settings['graph_url'] + '/me/photo/$value')

        if response.status_code == 200:
            picture_binary = response.content
            picture = "data:image/png;base64," + base64.b64encode(
                picture_binary).decode()  # builds the HTML src to display the base64 encoded image

        else:
            picture = None

        return picture

    def store_token(self, token):
        user_id = aa_users.get_current_user_id()
        aa_globals.session['microsoft_auth_token'] = token
        token = json.dumps({
            "token_type": token.get('token_type'),
            "scope": token.get('scope'),
            "expires_in": token.get('expires_in'),
            "ext_expires_in": token.get('ext_expires_in'),
            "access_token": token.get('access_token'),
            "refresh_token": token.get('refresh_token'),
            "id_token": token.get('id_token'),
            "expires_at": token.get('expires_at')
        })
        aa_database_queries.insert_provider_token(user_id=user_id, token=token, provider='microsoft')

    def store_user(self, user):
        aa_globals.session['user_email'] = user.get('mail') or user.get('userPrincipalName')
        aa_globals.session['auth_provider'] = 'microsoft'

    def callback(self, request):
        # Get the state saved in session
        expected_state = aa_globals.session.pop('microsoft_auth_state', '')
        # Make the token request
        token = self.get_token_from_code(request.url, expected_state)
        aa_globals.session['microsoft_auth_token'] = token

        user = self.get_user()
        self.store_user(user)

        log.info(f"User logged-in with microsoft account: {aa_globals.session.get('user_email')}")
