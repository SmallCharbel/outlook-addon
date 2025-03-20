// taskpane.js

Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // Initialize your add-in here
        console.log("PTA Metrics Approver Add-in is ready!");
        
        // Display current user email
        document.getElementById("current-email").textContent = Office.context.mailbox.userProfile.emailAddress;
        
        // Attach event handlers
        document.getElementById("sendMessageBtn").onclick = sendMessage;
        document.getElementById("sendMeetingRequestBtn").onclick = sendMeetingRequest;
    }
});

function showMessage(message, type) {
    var statusContainer = document.getElementById("status-container");
    statusContainer.innerHTML = message;
    statusContainer.className = type || "";
}

function sendMessage() {
    showMessage("Processing email...", "processing");
    
    // Get the current item
    var item = Office.context.mailbox.item;
    
    // Get the body first
    item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
        if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
            var htmlBody = bodyResult.value;
            
            // Get recipients
            item.getAsync(["to", "cc"], function(recipientsResult) {
                if (recipientsResult.status === Office.AsyncResultStatus.Succeeded) {
                    var toRecipients = recipientsResult.value.to || [];
                    var ccRecipients = recipientsResult.value.cc || [];
                    
                    // Create and send the new email
                    Office.context.mailbox.displayNewMessageForm({
                        toRecipients: toRecipients,
                        ccRecipients: ccRecipients,
                        subject: item.subject,
                        htmlBody: htmlBody
                    });
                    
                    // Move the original to deleted items
                    moveToDeletedItems();
                    
                    showMessage("Email forwarded successfully! Original will be moved to Deleted Items.", "success");
                } else {
                    showMessage("Error getting recipients: " + recipientsResult.error.message, "error");
                }
            });
        } else {
            showMessage("Error getting email body: " + bodyResult.error.message, "error");
        }
    });
}

function sendMeetingRequest() {
    showMessage("Processing meeting request...", "processing");
    
    // Get the current item
    var item = Office.context.mailbox.item;
    
    // Create a new appointment using the compose API
    Office.context.mailbox.displayNewAppointmentForm({
        requiredAttendees: getRecipientsArray(item.requiredAttendees),
        optionalAttendees: getRecipientsArray(item.optionalAttendees),
        subject: item.subject,
        body: item.body,
        start: item.start,
        end: item.end,
        location: item.location
    });
    
    // Move the original to deleted items
    moveToDeletedItems();
    
    showMessage("Meeting request sent successfully and original moved to Deleted Items.", "success");
}

// Convert callback-based APIs to Promises
function getRecipientsPromise(recipients) {
    return new Promise(function(resolve, reject) {
        if (!recipients) {
            resolve([]);
            return;
        }
        
        recipients.getAsync(function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var recipientArray = result.value.map(function(recipient) {
                    return {
                        displayName: recipient.displayName,
                        emailAddress: recipient.emailAddress
                    };
                });
                resolve(recipientArray);
            } else {
                reject(new Error("Failed to get recipients: " + result.error.message));
            }
        });
    });
}

function getBodyPromise(item) {
    return new Promise(function(resolve, reject) {
        item.body.getAsync(Office.CoercionType.Html, function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(new Error("Failed to get body: " + result.error.message));
            }
        });
    });
}

function getAttachmentsPromise(item) {
    return new Promise(function(resolve, reject) {
        if (!item.attachments || item.attachments.length === 0) {
            resolve([]);
            return;
        }
        
        item.attachments.getAsync(function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var attachmentArray = result.value.map(function(attachment) {
                    return {
                        type: "file",
                        name: attachment.name,
                        url: attachment.url,
                        isInline: false
                    };
                });
                resolve(attachmentArray);
            } else {
                reject(new Error("Failed to get attachments: " + result.error.message));
            }
        });
    });
}

function getRecipientsArray(recipients) {
    if (!recipients) return [];
    
    // Convert recipients to array format expected by displayNewMessageForm
    var recipientsArray = [];
    
    // Get recipients asynchronously
    recipients.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            recipientsArray = asyncResult.value.map(function(recipient) {
                return { 
                    displayName: recipient.displayName,
                    emailAddress: recipient.emailAddress 
                };
            });
        }
    });
    
    return recipientsArray;
}

function getAttachmentsArray(attachments) {
    if (!attachments) return [];
    
    // Convert attachments to array format expected by displayNewMessageForm
    var attachmentsArray = [];
    
    // Get attachments asynchronously
    attachments.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            attachmentsArray = asyncResult.value.map(function(attachment) {
                return {
                    id: attachment.id,
                    name: attachment.name,
                    size: attachment.size,
                    contentType: attachment.contentType
                };
            });
        }
    });
    
    return attachmentsArray;
}

function moveToDeletedItems() {
    if (Office.context.mailbox.item.move) {
        Office.context.mailbox.item.move(Office.MailboxEnums.FolderType.DeletedItems, {
            asyncContext: null,
            callback: function(asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showMessage("Failed to move item to Deleted Items: " + asyncResult.error.message, "error");
                } else {
                    console.log("Item moved to Deleted Items folder");
                }
            }
        });
    } else {
        console.log("Move API is not supported in this version.");
        showMessage("Move API is not supported in this version.", "warning");
    }
}
