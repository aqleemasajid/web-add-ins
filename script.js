Office.onReady(function() {
    // Initialization code
      // Get the current Word document
      var document = Office.context.document;
  
      // Load the properties you want to access
      document.load('properties/title, properties/author, properties/lastModifiedBy, properties/lastModifiedTime', function(result) {
        // Check if the load was successful
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Display the document information in the add-in
          var info = "Title: " + document.properties.title +
                     "\nAuthor: " + document.properties.author +
                     "\nLast Modified By: " + document.properties.lastModifiedBy +
                     "\nLast Modified Time: " + document.properties.lastModifiedTime;
          document.getElementById("content").innerHTML = "<pre>" + info + "</pre>";
        } else {
          console.log("Error loading document properties: " + result.error.message);
        }
      });
  
  });
  
