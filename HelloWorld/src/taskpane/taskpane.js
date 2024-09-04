Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("uploadButton").onclick = handleFileUpload;
  }
});

function handleFileUpload() {
  const fileInput = document.getElementById("emailFileInput");
  const file = fileInput.files[0];
  const messageDiv = document.getElementById("message");

  if (file) {
      const reader = new FileReader();

      reader.onload = function(event) {
          const emailData = event.target.result;
          messageDiv.textContent = "Email uploaded successfully!";
          console.log("Email Data:", emailData); // Log the email data to the console
      };

      reader.onerror = function() {
          messageDiv.textContent = "Error reading the file.";
      };

      reader.readAsText(file);
  } else {
      messageDiv.textContent = "Please select a file.";
  }
}
