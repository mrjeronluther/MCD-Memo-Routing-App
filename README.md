/**
 * Example: Triggering a Document Submission
 * This snippet demonstrates how the Vue.js frontend sends the 
 * sanitized payload to the processFormData Google Script function.
 */

const submitToMCDSystem = (payload) => {
  console.log("Initiating submission for:", payload.additionalMemoInfo);

  // Calling the Google Apps Script Backend
  google.script.run
    .withSuccessHandler((result) => {
      // result contains the generated Reference Number
      alert(`Success! Document logged as: ${result.refNumber}`);
      console.log("Confirmation details:", result);
    })
    .withFailureHandler((error) => {
      console.error("Submission failed:", error.message);
      alert("Error: " + error.message);
    })
    .processFormData(payload); 
    // 'processFormData' is the .gs function that handles File Upload & Sheet Logging
};

// Example Payload Structure
const samplePayload = {
  name: "John Doe",
  email: "jdoe@megaworld-lifestyle.com",
  memoType: "Jane Smith",
  dateType: "2025-06-15",
  additionalMemoInfo: "Request for Budget Realignment",
  selectedCGM: ["CAMSY"],
  fileData: "BASE64_ENCODED_STRING_HERE",
  fileMimeType: "application/pdf",
  fileName: "Budget_Memo.pdf"
};
