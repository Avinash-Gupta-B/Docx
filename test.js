const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
  path.resolve(__dirname, "template.docx"),
  "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip);

// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render({
  email: "user.email",
  name: "Avinash",
  designation: "IT",
  contactNumber: 8919,
  overview: `mznfkljd ffgmdkfg`,
  summary: "summary",
  summary1: "user.OverviewAndBilling.summary[1]",
  summary2: "user.OverviewAndBilling.summary[2]",
  summary3: "user.OverviewAndBilling.summary[3]",
  qualificationName: "B.Tech",
  qualificationUniversity: "JNTUK",
  skills: `I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.`,
  database: `Mongo DB`,
  frameworks: "Angular",
  git: "GIT , Github , Gitlab",
  platform: "VSCODE , Eclipse",
  projects: [
    {
      projectname: "qqqqqqqqq",
      projectdescription:
        "I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.",
      project1duration: "2023",
      projectrole: "Dev",
      toolsandtech: "zaxscsdvs",
      responsibility: "Skill",
    },
    {
      projectname: "qqqqqqqqq",
      projectdescription:
        "I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.",
      project1duration: "2023",
      projectrole: "Dev",
      toolsandtech: "zaxscsdvs",
      responsibility: "Skill",
    },
    {
      projectname: "qqqqqqqqq",
      projectdescription:
        "I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.I hereby declare that the above-mentioned information is correct up to my knowledge and I bear the responsibility for the correctness of the above-mentioned particulars.",
      project1duration: "2023",
      projectrole: "Dev",
      toolsandtech: "zaxscsdvs",
      responsibility: "Skill",
    },
  ],
});

const buf = doc.getZip().generate({
  type: "nodebuffer",
  // compression: DEFLATE adds a compression step.
  // For a 50MB output document, expect 500ms additional CPU time
  compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a
// file or res.send it with express for example.
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
