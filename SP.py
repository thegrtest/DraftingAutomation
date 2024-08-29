from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.listitem import ListItem
from office365.runtime.auth.client_credential import ClientCredential

# SharePoint site and credentials
site_url = "https://yourcompany.sharepoint.com/sites/yoursite"
client_id = "your-client-id"
client_secret = "your-client-secret"

# List name where the data will be written
list_name = "YourListName"

# Initialize the client context
credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(site_url).with_credentials(credentials)

# Function to write data to SharePoint list
def write_to_sharepoint(data):
    target_list = ctx.web.lists.get_by_title(list_name)
    item_creation_info = ListItem.create_entry(data)
    item = target_list.add_item(item_creation_info)
    ctx.execute_query()

    print(f"Item created: {item.id}")

# Example data to be written to the SharePoint list
data_to_write = {
    "Title": "New Item",  # Replace with the correct field name and value
    "CustomField": "Some value"  # Replace with the correct field name and value
}

# Write data to the SharePoint list
write_to_sharepoint(data_to_write)
