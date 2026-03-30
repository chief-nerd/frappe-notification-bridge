from . import __version__ as app_version

app_name = "frappe_notification_bridge"
app_title = "Frappe Notification Bridge"
app_publisher = "Mimirio"
app_description = "Bridge Frappe notifications to MS Teams chat."
app_email = "hello@mimirio.com"
app_license = "mit"

doc_events = {
	"Notification Log": {
		"after_insert": "frappe_notification_bridge.bridge.notify_teams"
	}
}
