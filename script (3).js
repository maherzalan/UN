// Function to get 'id' parameter from the URL
function getIdFromUrl() {
  const params = new URLSearchParams(window.location.search);
  return params.get("id");
}

// Function to process the Excel file
async function processExcelFile() {
  const id = getIdFromUrl();

  if (!id) {
    alert("No 'id' parameter found in the URL.");
    return;
  }

  try {
    // Fetch the Excel file from the project folder
    const response = await fetch("./data.xlsx");
    const arrayBuffer = await response.arrayBuffer();

    // Read the Excel file using SheetJS
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // Assuming the data is in the first sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert the sheet data to JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    // Find the row with the matching 'id'
    const matchingRow = jsonData.find((row) => row.Column1 == id);

    if (matchingRow) {
      // Redirect to the WhatsApp link
      window.location.href = matchingRow.whatsAppLink;
    } else {
      alert("No matching ID found in the Excel file.");
    }
  } catch (error) {
    console.error("Error processing the Excel file:", error);
    alert("Failed to read the Excel file.");
  }
}

// Call the function to process the Excel file
processExcelFile();
