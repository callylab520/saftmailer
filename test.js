const { createReport } = require('docx-templates')
const fs = require('fs')
const path = require("path");

require('dotenv').config()
const { MailtrapClient } = require("mailtrap");

const TOKEN = process.env.MAILER_TOKEN;
const ENDPOINT = process.env.MAILER_SERVER;
const client = new MailtrapClient({ endpoint: ENDPOINT, token: TOKEN });
const sender = {
  email: process.env.MAILER_SENDER,
  name: process.env.MAILER_NAME,
};

const sendEmail = async(receivers, subject, content, category) => {
  try {
    var recipients = []
    for (var x in receivers) {
      recipients.push({ email: receivers[x] })
    }

    const saftfile = fs.readFileSync(path.join(__dirname, "output.docx"));

    var ret = await client.send({
      from: sender,
      to: recipients,
      subject: subject,
      html: content,
      category: category,
      attachments: [
        { 
          filename: "saft.docx",
          content_id: "saft.docx",
          disposition: "inline",
          content: saftfile
        }
      ]
    })
    return ret
  } catch (error) {
    console.log("sendEmail", error)
    return {
      success: false, error: [ 'connection failed 2' ]
    }
  }
}
module.exports.sendEmail = sendEmail

const run = async() => {
  const template = fs.readFileSync('input.docx');

  const buffer = await createReport({
    template,
    data: {
      preparedfor: "John2",
      date: "4th August 2023",
    },
    additionalJsContext: {
      injectSvg: () => {
        const svg_data = Buffer.from(`<svg  xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
                                    <rect x="10" y="5" height="20" width="100" style="stroke:#ff0000; fill: #0000ff"/>
                                  </svg>`, 'utf-8');
  
        // Providing a thumbnail is technically optional, as newer versions of Word will just ignore it.
        const thumbnail = {
          data: fs.readFileSync('image.png'),
          extension: '.png',
        };
        return { width: 6, height: 3, data: svg_data, extension: '.svg', thumbnail };                    
      }
    },
    cmdDelimiter: ['{', '}'],
  });

  fs.writeFileSync('output.docx', buffer)

  var ret = await sendEmail(['sceptre520@gmail.com'], "TestSaft", "TestSaft Content", "Service");
  console.log(ret)
}
run()
