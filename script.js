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
    getAllAttachments();
}


async function getAllAttachments() {
    Office.context.mailbox.item?.getAttachmentsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error("Error getting all attachments");
    } else {
    console.info(JSON.stringify(asyncResult.value));
    asyncResult.value.map((i) => {
    ProcessAttachments(i.id);
    });
    }
    });
    }
    
    async function ProcessAttachments(id) {
    Office.context.mailbox.item?.getAttachmentContentAsync(id, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error(`Error getting attachment content for ${id}`);
    } else {
    console.info(`Got attachment content for ${id}`);
    }
    });
    }

