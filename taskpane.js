// taskpane.js

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sendMessageBtn").onclick = forwardEmail;
        
        // Display current user email
        document.getElementById("current-email").textContent = Office.context.mailbox.userProfile.emailAddress;
        
        // Get user initials for avatar
        const email = Office.context.mailbox.userProfile.emailAddress;
        if (email) {
            const nameParts = email.split('@')[0].split('.');
            let initials = "";
            if (nameParts.length >= 2) {
                initials = (nameParts[0].charAt(0) + nameParts[1].charAt(0)).toUpperCase();
            } else {
                initials = email.substring(0, 2).toUpperCase();
            }
            document.getElementById("user-initials").textContent = initials;
        }
    }
});

function updateStatus(message, type) {
    const statusContainer = document.getElementById("status-container");
    statusContainer.innerHTML = message;
    statusContainer.className = type || "";
}

function forwardEmail() {
    updateStatus("Processing email...", "processing");
    
    try {
        // Get the current item
        const item = Office.context.mailbox.item;
        
        // First, get the body content asynchronously
        item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                const htmlBody = bodyResult.value;
                
                // Create a completely new message with the same content
                const newMessageOptions = {
                    toRecipients: item.to,
                    ccRecipients: item.cc,
                    subject: item.subject,
                    htmlBody: htmlBody
                };
                
                // If there are attachments, try to include them
                if (item.attachments && item.attachments.length > 0) {
                    // Unfortunately, we can't directly copy attachments in the JavaScript API
                    // We'll need to notify the user about this limitation
                    updateStatus("Creating new email. Note: Attachments must be added manually.", "warning");
                }
                
                // Create the new message
                Office.context.mailbox.displayNewMessageForm(newMessageOptions);
                
                // After a short delay, try to move the original to deleted items
                setTimeout(() => {
                    moveToDeletedItems();
                }, 2000);
            } else {
                // If we can't get the body, still create the message but without body content
                Office.context.mailbox.displayNewMessageForm({
                    toRecipients: item.to,
                    ccRecipients: item.cc,
                    subject: item.subject
                });
                
                updateStatus("New email created (without body content). Please review and send.", "warning");
                
                // Still try to move the original
                setTimeout(() => {
                    moveToDeletedItems();
                }, 2000);
            }
        });
    } catch (error) {
        updateStatus(`Error: ${error.message}`, "error");
        console.error("Error creating new email:", error);
    }
}

function moveToDeletedItems() {
    try {
        const item = Office.context.mailbox.item;
        
        if (item.move) {
            item.move(Office.MailboxEnums.FolderType.DeletedItems, {
                asyncContext: null,
                callback: (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to move item:", result.error.message);
                    } else {
                        console.log("Original email moved to Deleted Items");
                        updateStatus("Original email moved to Deleted Items. Please review and send the new email.", "success");
                    }
                }
            });
        } else {
            console.log("Move API not supported in this version of Outlook");
            updateStatus("Move API not supported. Please delete the original email manually.", "warning");
        }
    } catch (error) {
        console.error("Error moving email to deleted items:", error);
        updateStatus("Error moving original email: " + error.message, "error");
    }
}
