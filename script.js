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
        var blob = new Blob([content], { type: 'application/octet-stream' });
        sendFileToJCT(blob);

        // Create an Object URL for the Blob
        var url = URL.createObjectURL(blob);

        // Create a download link
        var a = document.createElement('a');
        a.href = url;
        a.download = 'document.docx'; // Specify the desired file name
        a.style.display = 'none';

        // Append the link to the document body and trigger the download
        document.body.appendChild(a);
        a.click();

        // Clean up by revoking the Object URL
        URL.revokeObjectURL(url);
    }
}

function sendFileToJCT(fileData) {
    const endpointUrl = 'https://reqres.in/api/users'; // Replace with the actual URL

    fetch(endpointUrl, {
        method: 'POST',
        body: fileData, // The Word file data as a Blob or ArrayBuffer
        headers: {
            'Content-Type': 'application/octet-stream', // Adjust content type as needed
        },
    })
    .then(response => {
        if (response.ok) {
            console.log('File sent to JCT successfully.');
            var hasFocus = true;
            window.onblur = ()=> {
                hasFocus = false;
                window.onblur = null;
            };
            window.location.href = url;
            setTimeout(() => {
    
            }, this._onBlurWaitTime);
        } else {
            console.error('Failed to send file to JCT.');
        }
    })
    .catch(error => {
        console.error('Error sending file to JCT:', error);
    });
}


function myEncodeBase64(docData) {
    var s = "";
    for (var i = 0; i < docData.length; i++)
        s += String.fromCharCode(docData[i]);
    return window.btoa(s);
}
function sendSlice(slice, state) {
    var data = slice.data;

    if (data) {
        var fileData = myEncodeBase64(data);

        console.log(fileData);

        var request = new XMLHttpRequest();

        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                //updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                } else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "https://reqres.in/api/users");
        //request.setRequestHeader("Slice-Number", slice.index);
        updateStatus("File content: " + data);
        request.send(data);
    }
}

function closeFile(state) {
    state.file.closeAsync(function (result) {

        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus("File closed.");
        } else {
            updateStatus("File couldn't be closed.");
        }
    });
}
