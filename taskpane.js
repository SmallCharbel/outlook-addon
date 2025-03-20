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
        
        // Get recipients and subject
        let toRecipients = "";
        let ccRecipients = "";
        
        // Get the recipients from the current item
        if (item.to) {
            toRecipients = item.to;
        }
        
        if (item.cc) {
            ccRecipients = item.cc;
        }
        
        // Create a new message with the same recipients and subject
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: toRecipients,
            ccRecipients: ccRecipients,
            subject: item.subject
        });
        
        // After a short delay, try to move the original to deleted items
        setTimeout(() => {
            moveToDeletedItems();
        }, 2000);
        
        updateStatus("New email created! Please review and send.", "success");
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
                    }
                }
            });
        } else {
            console.log("Move API not supported in this version of Outlook");
        }
    } catch (error) {
        console.error("Error moving email to deleted items:", error);
    }
}
