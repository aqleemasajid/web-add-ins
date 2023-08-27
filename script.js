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
    console.log('here');
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
}

function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        } else {
            updateStatus(result.status);
        }
    });
}
function myEncodeBase64(docData)
{
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
