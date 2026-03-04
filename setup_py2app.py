from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)
