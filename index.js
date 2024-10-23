Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("insertSignature").onclick = function () {
            insertSignature();
        };
    }
});

function insertSignature() {
    const signature = `
        <br><br>Best Regards,<br><b>Your Name</b><br>Your Company<br>Your Position<br>
        <a href="https://Google.com">google.com</a>
    `;

    Office.context.mailbox.item.body.setAsync(
        signature,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Signature inserted successfully.");
            } else {
                console.error("Error inserting signature: " + asyncResult.error.message);
            }
        }
    );
}
