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
  loadSettings();
  if (info.host === Office.HostType.Outlook) {
    loadEmailContent();
    document.getElementById("saveBtn").addEventListener("click", () => {
      const noteTitle = document.getElementById("frontMatter").value.match(/title: "(.*)"/)[1];
      const sanitizedNoteTitle = sanitizeTitle(noteTitle);
      const yamlContent = document.getElementById("frontMatter").value + "\n";
      const markdownContent = emailContentEditor.value();

      saveToObsidian(sanitizedNoteTitle, yamlContent, markdownContent);
    });

  }
});

// Initialize the markdown editor
const emailContentEditor = new EasyMDE({maxHeight: "500px", element: document.getElementById("emailContent") });


async function loadEmailContent() {
  // Get the selected message in Outlook
  const item = Office.context.mailbox.item;
  console.log(item);
  const settings = Office.context.roamingSettings;
  const defaultTags = settings.get('defaultTags') || ''; // Replace with your default tags
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
  const markdownBody = sanitizeMarkdown(turndownService.turndown(emailBody));
  
  // Add horizontal lines to the markdown body whenever there is a new email
  // const markdownBody = turndownService.turndown(emailBody);

  // Fetch the attachments
  const attachments = await getAttachments(item);

  // Generate the default front matter
  const frontMatter = {
    title: subject,
    "date created": formattedDate,
    sender: fromName,
    "sender email": from,
    tags: defaultTags.split(",").map((tag) => tag.trim()),
  };

  // Update the UI with the generated content
  document.getElementById("frontMatter").value = dumpToFrontMatter(frontMatter);

  generateEditor(markdownBody, attachments);
  // Add attachments to the UI
  await renderAttachments(markdownBody, attachments);
}

function generateEditor(markdownBody, attachments) {
  // Generate the attachment links for the markdown note
  const attachmentLinks = generateAttachmentLinks(attachments);

  // Update the markdown body with the attachment links
  const markdownBodyWithAttachments = markdownBody + attachmentLinks;


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


// Replace this line:
// alert("Email content saved as a note in Obsidian.");
// With:

function sanitizeMarkdown(markdown) {
  // // Remove everything before the banner
  // markdown = markdown.replace(/[\s\S]*ZjQcmQRYFpfptBannerEnd\s*/, '');
  const settings = Office.context.roamingSettings;
  const attachmentFolder = settings.get('attachmentFolder') || 'Attachments'; 
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

  // Insert a horizontal line between emails
  markdown = markdown.replace(/\*\*From:\*\*/g, '\n---\n\n**From**');

  markdown = markdown.replace(/!\[\]\(cid:(image\d+\.(png|jpg|jpeg|gif|bmp|tiff))@[\w\d.]+\)/g, (match, imageName) => {
    return `![${imageName}](${attachmentFolder}/${imageName})`;
  });

  return markdown;
}

function dumpToFrontMatter(data) {
  const frontMatterData = { ...data, title: undefined }; // Exclude subject from yaml.dump
  let frontMatter = yaml.dump(frontMatterData);
  frontMatter = `title: "${data.title}"\n` + frontMatter; // Add subject as title separately
  console.log(frontMatter);
  return frontMatter;
}

function generateAttachmentLinks(attachments) {
  const settings = Office.context.roamingSettings;
  const attachmentFolder = settings.get('attachmentFolder') || 'Attachments'; // Replace with your default folder name
  let attachmentLinks = "\n\n## Attachments\n\n";

  if (attachments.length === 0) {
    return "";
  }

  attachments.forEach((attachment) => {
    attachment.contentUrl = attachmentFolder + "/" + encodeURIComponent(attachment.name);
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
  // Convert base64 to binary
  const byteCharacters = atob(attachment.content);
  const byteNumbers = new Array(byteCharacters.length);

  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }

  const byteArray = new Uint8Array(byteNumbers);
  const attachmentBlob = new Blob([byteArray], { type: attachment.contentType });

  // Save the attachment using the File System Access API
  const fileHandle = await window.showSaveFilePicker({
    suggestedName: attachment.name,
    // startIn: 'documents', // Start in the user's Documents folder
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
function renderAttachments(markdownBody, attachments) {
  const attachmentTableBody = document.getElementById('attachments-list');

  // Clear the attachment table body
  attachmentTableBody.innerHTML = '';

  attachments.forEach((attachment) => {
    const tableRow = document.createElement('tr');

    // Attachment name
    const nameCell = document.createElement('td');
    nameCell.className = 'py-2 text-left px-4 border-t border-l border-b';
    const attachmentName = document.createTextNode(attachment.name);
    nameCell.appendChild(attachmentName);
    tableRow.appendChild(nameCell);

    // Delete button
    const deleteCell = document.createElement('td');
    deleteCell.className = 'px-4 py-2 text-center border-t border-b';
    const deleteButton = document.createElement('button');
    deleteButton.className =
      'inline-block px-6 py-2 text-lg text-center text-white font-bold bg-blue-500 hover:bg-blue-600 focus:ring-4 focus:ring-blue-200 rounded-full';
    deleteButton.innerText = 'Delete';
    deleteButton.onclick = () => {
      attachments = removeAttachment(attachments, attachment.id);
      generateEditor(markdownBody, attachments);
      renderAttachments(markdownBody, attachments);
    };
    deleteCell.appendChild(deleteButton);
    tableRow.appendChild(deleteCell);

    // Save button
    const saveCell = document.createElement('td');
    saveCell.className = 'px-4 py-2 text-center border-t border-b border-r';
    const saveButton = document.createElement('button');
    saveButton.className =
      'inline-block px-6 py-2 text-lg text-center text-white font-bold bg-blue-500 hover:bg-blue-600 focus:ring-4 focus:ring-blue-200 rounded-full';
    saveButton.innerText = 'Save';
    saveButton.onclick = () => saveAttachmentToVault(attachment);
    saveCell.appendChild(saveButton);
    tableRow.appendChild(saveCell);

    attachmentTableBody.appendChild(tableRow);
  });
}

// Function to save the note to Obsidian
function saveToObsidian(noteTitle, yamlContent, markdownContent) {
  const settings = Office.context.roamingSettings;
  const vaultName = settings.get('vaultName') || 'Moomba PSD Knowledge Vault'; // Replace with your vault name
  const folderName = settings.get('defaultFolder') || 'Inbox'; // Replace with your default folder name
  const path = folderName ? `${folderName}/${noteTitle}` : noteTitle; // If there is a folder name, add it to the path
  const encodedTitle = encodeURIComponent(path);
  const encodedYamlContent = encodeURIComponent("---\n" + yamlContent + "---\n");
  const encodedMarkdownContent = encodeURIComponent(markdownContent);

  const obsidianUri = `obsidian://new?vault=${vaultName}&file=${encodedTitle}&content=${encodedYamlContent}${encodedMarkdownContent}`;

  window.open(obsidianUri);
}

function sanitizeTitle(title) {
  const invalidCharacters = /[<>:"\/\\|?*]/g;
  return title.replace(invalidCharacters, '');
}

function loadSettings() {
  const settings = Office.context.roamingSettings;
  console.log(settings);
}