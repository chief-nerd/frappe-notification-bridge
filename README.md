# Frappe Notification Bridge

Bridge Frappe notifications to MS Teams chat notifications for 1:1 user interaction. This plugin intercepts system notifications (mentions, assignments, alerts) and sends them as Adaptive Cards to the user's MS Teams chat.

## Features

- **Automated Interception**: Hooks into `Notification Log` to catch all system-generated notifications.
- **SSO Mapping**: Automatically maps Frappe users to MS Teams users based on their `@mimirio.com` email address.
- **Rich Notifications**: Uses Microsoft Adaptive Cards for a polished look and feel.
- **One-Click Access**: Includes a "View in ERPNext" button that deep-links directly to the relevant document.

## Setup Instructions

### 1. Microsoft Azure Configuration

To use this bridge, you must register an application in the [Microsoft Entra ID (Azure AD) portal](https://portal.azure.com/).

1.  **Register a New App**: Go to "App registrations" > "New registration".
2.  **API Permissions**: Add the following **Application Permissions** (not Delegated) for Microsoft Graph:
    - `User.Read.All` (To find the user's MS ID by email)
    - `Chat.Create` (To initiate 1:1 chats)
    - `ChatMessage.Send` (To send the notifications)
3.  **Grant Admin Consent**: Ensure you click "Grant admin consent for your organization" for these permissions.
4.  **Create Client Secret**: Go to "Certificates & secrets" and create a new client secret. Save this value.

### 2. Plugin Installation

Install the plugin on your bench:

```bash
bench get-app https://github.com/chief-nerd/frappe-notification-bridge
bench --site [your-site] install-app frappe_notification_bridge
```

### 3. ERPNext Configuration

1.  Log in to ERPNext as a **System Manager**.
2.  Search for **MS Teams Bridge Settings** in the Awesomebar.
3.  Fill in the credentials obtained from Azure:
    - **Enabled**: Check this to start the bridge.
    - **Tenant ID**: Your Azure Tenant ID.
    - **Client ID**: Your Azure Application (client) ID.
    - **Client Secret**: The secret value you generated.
4.  **Save**.

## How it works

The bridge listens for new entries in the `Notification Log`. If a notification is created for a user with a `@mimirio.com` email, it:
1. Fetches an access token from Microsoft.
2. Resolves the user's Microsoft ID.
3. Creates or locates a 1:1 chat between the Bot and the User.
4. Posts an Adaptive Card containing the notification subject, content, and a direct link.
