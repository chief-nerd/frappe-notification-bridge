import frappe
import requests
import json
from frappe import _
from frappe.utils import get_url, get_url_to_form

def notify_teams(doc, method=None):
	"""
	Hook called after_insert of Notification Log.
	Sends the notification to MS Teams as a 1:1 chat.
	"""
	if not doc.for_user:
		return

	# Get MS Teams settings
	try:
		settings = frappe.get_single("MS Teams Bridge Settings")
	except frappe.DoesNotExistError:
		return

	if not settings.enabled:
		return

	# Get user's email
	user_email = frappe.db.get_value("User", doc.for_user, "email")
	if not user_email or not user_email.endswith("@mimirio.com"):
		# User is not a Mimirio user (or SSO), skipping for now as per requirements
		return

	# Logic to send notification
	try:
		send_to_teams(doc, user_email, settings)
	except Exception as e:
		frappe.log_error(f"MS Teams Notification Bridge Error: {str(e)}", "MS Teams Bridge")

def send_to_teams(doc, user_email, settings):
	token = get_access_token(settings)
	if not token:
		return

	# Get MS User ID
	ms_user_id = get_ms_user_id(user_email, token)
	if not ms_user_id:
		return

	# Get or create 1:1 chat ID
	chat_id = get_or_create_chat_id(ms_user_id, settings.client_id, token)
	if not chat_id:
		return

	# Prepare Adaptive Card
	card = prepare_adaptive_card(doc)

	# Send message
	url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
	headers = {
		"Authorization": f"Bearer {token}",
		"Content-Type": "application/json"
	}
	payload = {
		"body": {
			"contentType": "html",
			"content": '<attachment id="74d20c7f-3451-4091-9e79-563d4d42065b"></attachment>'
		},
		"attachments": [
			{
				"id": "74d20c7f-3451-4091-9e79-563d4d42065b",
				"contentType": "application/vnd.microsoft.card.adaptive",
				"content": json.dumps(card)
			}
		]
	}
	res = requests.post(url, headers=headers, json=payload)
	if res.status_code not in [200, 201]:
		frappe.log_error(f"Failed to send MS Teams message: {res.text}", "MS Teams Bridge")

def get_access_token(settings):
	url = f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token"
	data = {
		"client_id": settings.client_id,
		"scope": "https://graph.microsoft.com/.default",
		"client_secret": settings.get_password("client_secret"),
		"grant_type": "client_credentials"
	}
	res = requests.post(url, data=data)
	if res.status_code != 200:
		frappe.log_error(f"Failed to get MS Teams access token: {res.text}", "MS Teams Bridge")
		return None
	return res.json().get("access_token")

def get_ms_user_id(user_email, token):
	url = f"https://graph.microsoft.com/v1.0/users/{user_email}"
	headers = {"Authorization": f"Bearer {token}"}
	res = requests.get(url, headers=headers)
	if res.status_code != 200:
		return None
	return res.json().get("id")

def get_or_create_chat_id(ms_user_id, bot_app_id, token):
	url = "https://graph.microsoft.com/v1.0/chats"
	headers = {
		"Authorization": f"Bearer {token}",
		"Content-Type": "application/json"
	}
	payload = {
		"chatType": "oneOnOne",
		"members": [
			{
				"@odata.type": "#microsoft.graph.aadUserConversationMember",
				"roles": ["owner"],
				"user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{ms_user_id}')"
			},
			{
				"@odata.type": "#microsoft.graph.aadUserConversationMember",
				"roles": ["owner"],
				"user@odata.bind": f"https://graph.microsoft.com/v1.0/applications('{bot_app_id}')"
			}
		]
	}
	# Try to find existing chat first or create new
	# For simplicity in this bridge, we use the creation endpoint which returns existing if it exists
	res = requests.post(url, headers=headers, json=payload)
	if res.status_code in [200, 201]:
		return res.json().get("id")
	
	# If creation fails, it might be because the bot is an application and needs a different member type
	# or the bot is already in a chat. 
	# Fallback: list chats for the user (requires high permissions)
	frappe.log_error(f"Failed to create/get MS Teams chat: {res.text}", "MS Teams Bridge")
	return None

def prepare_adaptive_card(doc):
	link = doc.link
	if not link and doc.document_type and doc.document_name:
		link = get_url_to_form(doc.document_type, doc.document_name)
	
	if link and link.startswith("/"):
		link = get_url() + link

	# Strip HTML from email_content if present for text blocks
	from frappe.utils import strip_html
	content = strip_html(doc.email_content or "")

	card = {
		"type": "AdaptiveCard",
		"version": "1.4",
		"body": [
			{
				"type": "TextBlock",
				"text": doc.subject or _("New Notification"),
				"weight": "Bolder",
				"size": "Medium",
				"wrap": True
			},
			{
				"type": "TextBlock",
				"text": content,
				"wrap": True
			}
		],
		"actions": []
	}

	if link:
		card["actions"].append({
			"type": "Action.OpenUrl",
			"title": _("View in ERPNext"),
			"url": link
		})

	return card
