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
    Office.context.document.getFileAsync(Office.FileType.Text,
        { sliceSize: 100000 },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            } else {
                updateStatus(result.status);
            }
        });

    function getSlice(state) {
        state.file.getSliceAsync(state.counter, function (sliceResult) {
            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                // Concatenate the content slices
                state.content = (state.content || '') + sliceResult.value.data;

                // Check if there are more slices
                if (state.counter < state.sliceCount - 1) {
                    state.counter++;
                    getSlice(state);
                } else {
                    // All slices are retrieved, now save the content locally
                    saveContentLocally(state.content);
                }
            } else {
                updateStatus(sliceResult.status);
            }
        });
    }

    function saveContentLocally(content) {
        // Create a Blob with the document content
        var blob = new Blob([content], { type: 'text/plain' });

        // Create an Object URL for the Blob
        var url = URL.createObjectURL(blob);

        // Create a download link
        var a = document.createElement('a');
        a.href = url;
        a.download = 'document.txt'; // Specify the desired file name
        a.style.display = 'none';

        // Append the link to the document body and trigger the download
        document.body.appendChild(a);
        a.click();

        // Clean up by revoking the Object URL
        URL.revokeObjectURL(url);
    }
}
