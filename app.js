const XLSX = require("xlsx");
const axios = require("axios");
const path = require("path");
const { log } = require("console");

// Load Excel file
const workbook = XLSX.readFile(path.join(__dirname, "MERGED.xlsx"));
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert sheet to JSON
let rows = XLSX.utils.sheet_to_json(sheet);

rows = rows.filter((row) => row["Mobile No"] && row["Mobile No"].toString().length == 13);

console.log(`Total rows found: ${rows.length}`);

async function callApiForRow(row, index) {
  try {
    // const response = await axios.post(
    //   "https://example.com/api/users",
    //   {
    //     name: row.name,
    //     mobile: row.mobile,
    //     email: row.email,
    //   },
    //   {
    //     timeout: 10000,
    //   }
    // );

    // console.log(`Row ${index + 1} ✅ Success`, response.data);
    if (row["Mobile No"] && row["Mobile No"].toString().length == 13) {
      await sendSMS(row);
      // console.log(`Row ${index + 1} ✅ Success`, row["Mobile No"]?.substr(1));
    }
  } catch (error) {
    console.error(`Row ${index + 1} ❌ Failed`, error.response?.data || error.message);
  }
}

// Sequential processing (SAFE)
async function processExcel() {
  for (let i = 0; i < rows.length; i++) {
    if (rows[i]["Mobile No"] && rows[i]["Mobile No"].toString().length == 13) {
      await callApiForRow(rows[i], i);
      await sleep(800);
    }
  }

  console.log("🎉 Excel processing completed");
}

processExcel();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function sendSMS(row) {
  let data = JSON.stringify({
    messaging_product: "whatsapp",
    to: row["Mobile No"]?.substr(1),
    type: "template",
    template: {
      name: "accout_update_for_new_users",
      language: {
        code: "en",
      },
      components: [
        {
          type: "header",
          parameters: [
            {
              type: "image",
              image: {
                link: "https://learnassets.sk-coders.com/emps/2026/b1.jpeg",
              },
            },
          ],
        },
        {
          type: "body",
          parameters: [
            {
              type: "text",
              text: row["Sr No"]?.toString()?.trim() || "-",
            },
            {
              type: "text",
              text: row["Full Name (English)"]?.toString()?.trim() || "-",
            },
            {
              type: "text",
              text: row["Voter ID (EPIC)"]?.toString()?.trim() || "-",
            },
            {
              type: "text",
              text: row["Booth Address (Marathi)"]?.toString()?.trim() || "-",
            },
          ],
        },
      ],
    },
  });

  let config = {
    method: "post",
    maxBodyLength: Infinity,
    url: "https://graph.facebook.com/v22.0/721104671093303/messages",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer EAAK4ib0AF0sBQf2XeuIAYlSU4vgjlOOtItNLBiW5w25ZCtvvYOTtutHvosLgDkQ4ZCwEPR76fLp6Uvjz8rmaXj4n0LkGmjrAQIkPZCEWa9OkMWoHAQNZBJepuWZAh5IfnIFBZBYBVoMgoM6umIkBtwZAjBQqPbA5esHWo56Q96AZBZA4i36SgC3QnLig4e1h3dAC3L7VOASR4FQBKJh6DOFGzYrc1OEgTuMzRrN9Y",
    },
    data: data,
  };
  try {
    const res = await axios.request(config);
    console.log("MSG SENT On :" + row["Mobile No"]);
  } catch (err) {
    console.log("Error in sending SMS", err);
  }
}
