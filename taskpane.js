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
                
                // Check if there are attachments
                if (item.attachments && item.attachments.length > 0) {
                    // Get attachments
                    item.getAttachmentsAsync((attachmentsResult) => {
                        if (attachmentsResult.status === Office.AsyncResultStatus.Succeeded) {
                            const attachments = attachmentsResult.value;
                            
                            // Prepare attachments for the new message
                            const attachmentArray = attachments.map(attachment => {
                                return {
                                    type: "file",
                                    name: attachment.name,
                                    url: attachment.url || attachment.id,
                                    isInline: false
                                };
                            });
                            
                            // Create new message with attachments
                            Office.context.mailbox.displayNewMessageForm({
                                toRecipients: item.to,
                                ccRecipients: item.cc,
                                subject: item.subject,
                                htmlBody: htmlBody,
                                attachments: attachmentArray
                            });
                            
                            // Move original to deleted items
                            setTimeout(() => {
                                moveToDeletedItems();
                            }, 2000);
                        } else {
                            // Failed to get attachments, create message without them
                            createMessageWithoutAttachments(item, htmlBody);
                        }
                    });
                } else {
                    // No attachments, create simple message
                    createMessageWithoutAttachments(item, htmlBody);
                }
            } else {
                // Failed to get body, create message with just subject and recipients
                Office.context.mailbox.displayNewMessageForm({
                    toRecipients: item.to,
                    ccRecipients: item.cc,
                    subject: item.subject
                });
                
                updateStatus("New email created (without body content). Please review and send.", "warning");
                
                // Move original
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

function createMessageWithoutAttachments(item, htmlBody) {
    Office.context.mailbox.displayNewMessageForm({
        toRecipients: item.to,
        ccRecipients: item.cc,
        subject: item.subject,
        htmlBody: htmlBody
    });
    
    updateStatus("New email created. Please review and send.", "success");
    
    // Move original
    setTimeout(() => {
        moveToDeletedItems();
    }, 2000);
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
