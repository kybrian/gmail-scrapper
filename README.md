# Gmail API Integration

## Step 1: Enable the Gmail API

### Go to the Google Cloud Console

Open your web browser and go to the [Google Cloud Console](https://console.cloud.google.com/).

### Create a New Project (If you don't have one)

If this is your first time using Google Cloud, you'll need to create a project. Click the project dropdown in the top left corner and select **NEW PROJECT**. Give your project a name (e.g., "Gmail Downloader").

### Enable the Gmail API

1. Click the menu icon in the top left corner (three horizontal lines).
2. Go to **APIs & Services** > **Library**.
3. Search for **Gmail API** and click on it.
4. Click **Enable**.

### Create Credentials

1. Go to **APIs & Services** > **Credentials**.
2. Click **Create Credentials** > **OAuth client ID**.
3. Choose **Desktop app** for the application type.
4. Give it a name (e.g., "Gmail Downloader App").
5. Click **Create**.

### Download the Credentials

You'll see a message saying "OAuth client created". Click **DOWNLOAD JSON**. This will download a file named `credentials.json`.

**IMPORTANT:** Keep this file in a safe location because it grants access to your Gmail. You will need to save it in the same folder as your Python script.

## Step 2: Install Necessary Libraries

You will need Python (v3.6+) installed on your system.

### Open your Terminal or Command Prompt

Install the required libraries:

```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
