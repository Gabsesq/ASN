import os

# Define TMP_PATH to use '/tmp' on Vercel or 'tmp/' for local testing
TMP_PATH = '/tmp' if os.environ.get('VERCEL') else 'tmp'
