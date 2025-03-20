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
    
    // Get recipients properly with callbacks
    getRecipients(item, function(toRecipients, ccRecipients) {
        // Get body
        item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                var htmlBody = bodyResult.value;
                
                // Create a new message using the compose API
                Office.context.mailbox.displayNewMessageForm({
                    toRecipients: toRecipients,
                    ccRecipients: ccRecipients,
                    subject: item.subject,
                    htmlBody: htmlBody
                });
                
                // Move the original to deleted items
                moveToDeletedItems();
                
                showMessage("Email sent successfully and original moved to Deleted Items.", "success");
            } else {
                showMessage("Failed to get email body: " + bodyResult.error.message, "error");
            }
        });
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

function getRecipients(item, callback) {
    var toRecipients = [];
    var ccRecipients = [];
    
    // Get TO recipients
    if (item.to) {
        item.to.getAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                toRecipients = asyncResult.value.map(function(recipient) {
                    return { 
                        displayName: recipient.displayName,
                        emailAddress: recipient.emailAddress 
                    };
                });
                
                // Get CC recipients
                if (item.cc) {
                    item.cc.getAsync(function(ccResult) {
                        if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                            ccRecipients = ccResult.value.map(function(recipient) {
                                return { 
                                    displayName: recipient.displayName,
                                    emailAddress: recipient.emailAddress 
                                };
                            });
                        }
                        callback(toRecipients, ccRecipients);
                    });
                } else {
                    callback(toRecipients, ccRecipients);
                }
            } else {
                showMessage("Failed to get recipients: " + asyncResult.error.message, "error");
                callback([], []);
            }
        });
    } else {
        callback([], []);
    }
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
