Office.initialize = function (reason) {

    $(document).ready(function () {
        $('#submit').click(function () {
            sendFile();
        });

        updateStatus("Ready to send file.");
    });
}

function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}

function sendFile() {
    Office.context.mailbox.item.attachments.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            var attachments = asyncResult.value;
            attachments.forEach(function (attachment) {
                attachment.getAttachmentContentAsync(function (contentAsyncResult) {
                    if (contentAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        var attachmentContent = contentAsyncResult.value.data;
                        // Process attachment content here
                        console.log(attachmentContent);
                    } else {
                        console.error("Error retrieving attachment content: " + contentAsyncResult.error.message);
                    }
                });
            });
        } else {
            console.error("Error retrieving attachments: " + asyncResult.error.message);
        }
    });
}

