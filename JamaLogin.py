import requests
from requests.auth import HTTPBasicAuth

def JamaLogin(basic_oauth, jama_username, jama_password, jama_base_url_v2):
    """
    Authenticates with the Jama Connect API and returns an authenticated requests.Session object.

    Args:
        basic_oauth (str): The authentication method, either 'basic' or 'oauth'.
        jama_username (str): The username or client ID for authentication.
        jama_password (str): The password or client secret for authentication.
        jama_base_url_v2 (str): The base URL for the Jama Connect instance.

    Returns:
        requests.Session: An authenticated session object for making subsequent API calls.
    
    Raises:
        requests.exceptions.HTTPError: If the authentication fails due to an HTTP error.
        Exception: If an unexpected error occurs during authentication.
    """
    print(f"\nAttempting to authenticate with Jama Connect using {basic_oauth.upper()}...")
    session = requests.Session()
    
    if basic_oauth == 'basic':
        auth = HTTPBasicAuth(jama_username, jama_password)
        session.auth = auth
    elif basic_oauth == 'oauth':
        client_id = jama_username
        client_secret = jama_password
        token_url = f"{jama_base_url_v2.rstrip('/')}/rest/oauth/token"
        
        try:
            # Get the access token using HTTP Basic Authentication
            token_data = {'grant_type': 'client_credentials'}
            response = requests.post(token_url, data=token_data, auth=HTTPBasicAuth(client_id, client_secret))
            response.raise_for_status()
            token = response.json().get('access_token')
            
            # Use the token for subsequent requests
            session.headers.update({"Authorization": f"Bearer {token}"})
            print("OAuth 2.0 authentication successful! 🎉")
        except requests.exceptions.HTTPError as e:
            raise requests.exceptions.HTTPError(f"OAuth 2.0 authentication failed. Please check your client ID and secret. Error: {e}")
        except Exception as e:
            raise Exception(f"An unexpected error occurred during OAuth authentication: {e}")
    else:
        raise ValueError("Invalid 'basic_oauth' value. Please use 'basic' or 'oauth'.")
    
    # Test authentication by making a simple API call
    test_url = f"{jama_base_url_v2.rstrip('/')}/projects"
    try:
        response = session.get(test_url)
        response.raise_for_status()
        print("Authentication successful! 🎉")
    except requests.exceptions.HTTPError as e:
        raise requests.exceptions.HTTPError(f"Authentication failed. Please check your credentials. Error: {e}")
    except Exception as e:
        raise Exception(f"An unexpected error occurred during test call: {e}")
    
    print("\nAuthentication complete. Ready to make API calls.")
    
    return session