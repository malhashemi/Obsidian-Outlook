/* eslint-disable no-useless-escape */
/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
// import './taskpane.css';
import TurndownService from "turndown";
import * as yaml from "js-yaml";
import EasyMDE from "easymde";

const Office = window.Office;
const turndownService = new TurndownService();

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("saveBtn").onclick = saveToObsidian;
    loadEmailContent();
  }
});

// Initialize the markdown editor
const emailContentEditor = new EasyMDE({ element: document.getElementById("emailContent") });


async function loadEmailContent() {
  // Get the selected message in Outlook
  const item = Office.context.mailbox.item;
  console.log(item);

  // Get the email's subject, sender, and received time
  const subject = item.normalizedSubject;
  console.log(subject);
  const from = item.from.emailAddress;
  const fromName = item.from.displayName;
  const receivedTime = item.dateTimeCreated.toJSON();

  // change date format to YYYY-MM-DD HH:MM:SS AM/PM
  const date = new Date(receivedTime);
  const formattedDate =
    date.getFullYear() +
    "-" +
    ("0" + (date.getMonth() + 1)).slice(-2) +
    "-" +
    ("0" + date.getDate()).slice(-2) +
    " " +
    ("0" + date.getHours()).slice(-2) +
    ":" +
    ("0" + date.getMinutes()).slice(-2) +
    ":" +
    ("0" + date.getSeconds()).slice(-2) +
    " " +
    (date.getHours() > 12 ? "PM" : "AM");
  // Fetch the email's HTML body
  const emailBody = await getEmailBody(item);

  // Convert the HTML body to markdown
  const markdownBody = removeHtmlAndCss(turndownService.turndown(emailBody));
  // const markdownBody = turndownService.turndown(emailBody);

  // Fetch the attachments
  const attachments = await getAttachments(item);

  // Add attachments to the UI
  await renderAttachments(attachments);

  // Generate the attachment links for the markdown note
  const attachmentLinks = generateAttachmentLinks(attachments);

  // Update the markdown body with the attachment links
  const markdownBodyWithAttachments = markdownBody + attachmentLinks;


  // Generate the default front matter
  const frontMatter = {
    title: subject,
    "date created": formattedDate,
    sender: fromName,
    "sender email": from,
    tags: [],
  };
  // Update the UI with the generated content
  document.getElementById("frontMatter").value = dumpToFrontMatter(frontMatter);
  emailContentEditor.value(markdownBodyWithAttachments);
  //document.getElementById("emailContent").value = markdownBody;
}

function getEmailBody(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, {}, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        console.error("Failed to get email body.", result.error);
        resolve("");
      }
    });
  });
}

function showMessage(message) {
  const messageElement = document.getElementById("message");
  messageElement.innerHTML = message;
  messageElement.style.display = "block";
}

function hideMessage() {
  const messageElement = document.getElementById("message");
  messageElement.style.display = "none";
}

// Replace this line:
// alert("Email content saved as a note in Obsidian.");
// With:

function saveToObsidian() {
  // TODO: Implement saving the edited front matter and markdown content to Obsidian
  showMessage("Email content saved as a note in Obsidian.");
  setTimeout(hideMessage, 3000);
}

function removeHtmlAndCss(markdown) {
  // // Remove everything before the banner
  // markdown = markdown.replace(/[\s\S]*ZjQcmQRYFpfptBannerEnd\s*/, '');
  const bannerStart = 'ZjQcmQRYFpfptBannerStart';
  const bannerEnd = 'ZjQcmQRYFpfptBannerEnd';

  const startIndex = markdown.indexOf(bannerStart);
  const endIndex = markdown.indexOf(bannerEnd);

  if (startIndex !== -1 && endIndex !== -1) {
    markdown = markdown.slice(0, startIndex) + markdown.slice(endIndex + bannerEnd.length);
  }
  // Remove HTML tags, CSS style blocks, and CSS selectors
  markdown = markdown.replace(/<[^>]*>|<\/[^>]*>|\s*<!--[\s\S]*?-->|\s*{[^}]*}|#(\w|\d)+[^\s{]*\s*|\s*\.\w+(\s*\:\w+){1,3}\s*/g, '');

  // Remove the comma immediately following the artifacts
  markdown = markdown.replace(/^,/, '');

  // Remove extra spaces and new lines after the banner
  markdown = markdown.replace(/^\s+/, '');

  return markdown;
}

function dumpToFrontMatter(data) {
  const frontMatterData = { ...data, title: undefined }; // Exclude subject from yaml.dump
  let frontMatter = yaml.dump(frontMatterData);
  frontMatter = `title: "${data.title}"\n` + frontMatter; // Add subject as title separately
  console.log(frontMatter);
  return frontMatter;
}

function displayAttachments(attachments) {
  const attachmentsList = document.getElementById("attachments-list");

  if (attachments.length === 0) {
    attachmentsList.innerHTML = "No attachments.";
  } else {
    attachmentsList.innerHTML = "";
    for (const attachment of attachments) {
      const listItem = document.createElement("li");

      const attachmentName = document.createElement("span");
      attachmentName.textContent = attachment.name;
      listItem.appendChild(attachmentName);

      const saveButton = document.createElement("button");
      saveButton.textContent = "Save";
      saveButton.addEventListener("click", () => {
        saveToObsidian(attachment);
      });
      listItem.appendChild(saveButton);

      attachmentsList.appendChild(listItem);
    }
  }
}


function generateAttachmentLinks(attachments) {
  let attachmentLinks = "\n\n## Attachments\n\n";

  if (attachments.length === 0) {
    return "";
  }

  attachments.forEach((attachment) => {
    attachment.contentUrl = "Attachments/" + attachment.name;
    attachmentLinks += `- [${attachment.name}](${attachment.contentUrl})\n`;
  });

  return attachmentLinks;
}

async function getAttachments(item) {
  const attachments = item.attachments;

  if (attachments.length === 0) {
    return [];
  }

  const fetchedAttachments = [];

  for (const attachment of attachments) {
    const content = await getAttachmentContent(item, attachment.id);
    fetchedAttachments.push({
      ...attachment,
      content: content.content,
      format: content.format,
    });
  }
  console.log(fetchedAttachments);
  return fetchedAttachments;
}

function getAttachmentContent(item, attachmentId) {
  return new Promise((resolve) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        console.error("Failed to get attachment content.", result.error);
        resolve(null);
      }
    });
  });
}

// Function to save an attachment in the Obsidian vault
async function saveAttachmentToVault(attachment) {
  const attachmentData = atob(attachment.content); // Convert base64 to binary
  const attachmentBlob = new Blob([attachmentData], { type: attachment.contentType });

  // Replace this with the path to your Obsidian vault's Attachments folder
  const vaultPath = '/path/to/your/obsidian/vault/Attachments/';

  // const filePath = `${vaultPath}${attachment.name}`;

  // Save the attachment to the specified path
  // This assumes you have the necessary permissions to write to the filesystem
  const fileHandle = await window.showSaveFilePicker({
    suggestedName: attachment.name,
    startIn: vaultPath,
  });

  const writable = await fileHandle.createWritable();
  await writable.write(attachmentBlob);
  await writable.close();
}

// Function to remove an attachment from the list
function removeAttachment(attachments, attachmentId) {
  return attachments.filter((attachment) => attachment.id !== attachmentId);
}

// Function to render attachments with Save and Delete buttons
function renderAttachments(attachments) {
  const attachmentList = document.getElementById('attachments-list');

  // Clear the attachment list
  attachmentList.innerHTML = '';

  attachments.forEach((attachment) => {
    const listItem = document.createElement('li');
    const attachmentName = document.createTextNode(attachment.name);
    listItem.appendChild(attachmentName);

    const saveButton = document.createElement('button');
    saveButton.innerText = 'Save';
    saveButton.onclick = () => saveAttachmentToVault(attachment);
    listItem.appendChild(saveButton);

    const deleteButton = document.createElement('button');
    deleteButton.innerText = 'Delete';
    deleteButton.onclick = () => {
      attachments = removeAttachment(attachments, attachment.id);
      renderAttachments(attachments);
    };
    listItem.appendChild(deleteButton);

    attachmentList.appendChild(listItem);
  });
}