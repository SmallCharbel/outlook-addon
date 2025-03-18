// taskpane.js

Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // Initialize your add-in here
        console.log("PTA Metrics Approver Add-in is ready!");
        
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
    
    Office.context.mailbox.item.to.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showMessage("Failed to get recipients: " + asyncResult.error.message, "error");
            return;
        }
        
        var recipTO = asyncResult.value.map(function(recipient) { 
            return recipient.emailAddress; 
        }).join(';');
        
        Office.context.mailbox.item.cc.getAsync(function (ccAsyncResult) {
            if (ccAsyncResult.status === Office.AsyncResultStatus.Failed) {
                showMessage("Failed to get CC recipients: " + ccAsyncResult.error.message, "error");
                return;
            }
            
            var recipCC = ccAsyncResult.value.map(function(recipient) { 
                return recipient.emailAddress; 
            }).join(';');
            
            Office.context.mailbox.item.subject.getAsync(function (subjectAsyncResult) {
                if (subjectAsyncResult.status === Office.AsyncResultStatus.Failed) {
                    showMessage("Failed to get subject: " + subjectAsyncResult.error.message, "error");
                    return;
                }
                
                var subject = subjectAsyncResult.value;
                
                Office.context.mailbox.item.body.getAsync('html', function (bodyAsyncResult) {
                    if (bodyAsyncResult.status === Office.AsyncResultStatus.Failed) {
                        showMessage("Failed to get body: " + bodyAsyncResult.error.message, "error");
                        return;
                    }
                    
                    var body = bodyAsyncResult.value;
                    
                    createAndSendEmail(recipTO, recipCC, subject, body);
                });
            });
        });
    });
}

function sendMeetingRequest() {
    showMessage("Processing meeting request...", "processing");
    
    Office.context.mailbox.item.to.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showMessage("Failed to get recipients: " + asyncResult.error.message, "error");
            return;
        }
        
        var recipTO = asyncResult.value.map(function(recipient) { 
            return recipient.emailAddress; 
        }).join(';');
        
        Office.context.mailbox.item.cc.getAsync(function (ccAsyncResult) {
            if (ccAsyncResult.status === Office.AsyncResultStatus.Failed) {
                showMessage("Failed to get CC recipients: " + ccAsyncResult.error.message, "error");
                return;
            }
            
            var recipCC = ccAsyncResult.value.map(function(recipient) { 
                return recipient.emailAddress; 
            }).join(';');
            
            Office.context.mailbox.item.subject.getAsync(function (subjectAsyncResult) {
                if (subjectAsyncResult.status === Office.AsyncResultStatus.Failed) {
                    showMessage("Failed to get subject: " + subjectAsyncResult.error.message, "error");
                    return;
                }
                
                var subject = subjectAsyncResult.value;
                
                Office.context.mailbox.item.body.getAsync('html', function (bodyAsyncResult) {
                    if (bodyAsyncResult.status === Office.AsyncResultStatus.Failed) {
                        showMessage("Failed to get body: " + bodyAsyncResult.error.message, "error");
                        return;
                    }
                    
                    var body = bodyAsyncResult.value;
                    
                    createAndSendMeetingRequest(recipTO, recipCC, subject, body);
                });
            });
        });
    });
}

function createAndSendEmail(toRecipients, ccRecipients, subject, body) {
    Office.context.mailbox.item.attachments.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showMessage("Failed to get attachments: " + asyncResult.error.message, "error");
            return;
        }
        
        var attachments = asyncResult.value;
        
        Office.context.mailbox.item.saveAsync(function (saveResult) {
            if (saveResult.status === Office.AsyncResultStatus.Failed) {
                showMessage("Failed to save item: " + saveResult.error.message, "error");
                return;
            }
            
            // Use Office.context.mailbox.displayReplyAllForm instead of replyAll()
            var replyHtml = body;
            
            // Create the reply details
            var replyDetails = {
                subject: subject,
                htmlBody: replyHtml
            };
            
            // If there are attachments, we need to handle them
            if (attachments && attachments.length > 0) {
                showMessage("Processing attachments...", "processing");
                
                // For this simplified version, we'll just send without attachments
                Office.context.mailbox.displayReplyAllForm(replyDetails);
                
                showMessage("Email sent. Attachments are not supported in this version.", "success");
            } else {
                // No attachments, send the email
                Office.context.mailbox.displayReplyAllForm(replyDetails);
                
                // Move the original to deleted items
                moveToDeletedItems();
                
                showMessage("Email sent successfully and original moved to Deleted Items.", "success");
            }
        });
    });
}

function createAndSendMeetingRequest(toRecipients, ccRecipients, subject, body) {
    Office.context.mailbox.item.attachments.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showMessage("Failed to get attachments: " + asyncResult.error.message, "error");
            return;
        }
        
        var attachments = asyncResult.value;
        
        showMessage("Meeting request functionality is not fully supported in this version.", "processing");
        
        Office.context.mailbox.item.saveAsync(function () {
            // Move the original to deleted items
            moveToDeletedItems();
            
            showMessage("Meeting request processed. Note: Full meeting request functionality requires additional implementation.", "success");
        });
    });
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
    }
}
