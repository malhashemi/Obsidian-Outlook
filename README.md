# Add Email as Obsidian Note Outlook Add-in

This is an Outlook add-in that allows users to save an email as a note in their Obsidian vault, complete with front matter and attachments.

## Features

- **Save Email as Obsidian Note**: Convert the email body to markdown and save it as a note in the Obsidian vault.
- **User Settings**: Specify the Obsidian vault to which the notes should be saved.
- **Attachments Handling**: Save attachments to a designated folder in the Obsidian vault and link them in the markdown note.

## Installation

This add-in needs to be hosted on a server with a valid SSL certificate. One option is to use Firebase Hosting.

Here's a guide on how to host this add-in using Firebase:

### Prerequisites

- Install [Node.js and npm](https://nodejs.org/en/download/)
- Install [Firebase CLI](https://firebase.google.com/docs/cli) by running `npm install -g firebase-tools`

### Steps

1. Clone this repository:

    ```sh
    git clone https://github.com/malhashemi/Obsidian-Outlook.git
    ```

2. Navigate to the project directory:

    ```sh
    cd Obsidian-Outlook
    ```

3. Install npm

    ```sh
    npm install
    ```

4. Login to Firebase:

    ```sh
    firebase login
    ```

5. Initialize Firebase in your project:

    ```sh
    firebase init
    ```

    - Select "Hosting"
    - Select "Create a new project" and follow the prompts to create a new project on Firebase
    - Specify "dist" as the public directory
    - Choose "Yes" to configure the project as a single-page app

6. Build the project:

    ```sh
    npm run build
    ```

    This will create a `dist` directory with the built project files.

7. Deploy the project to Firebase:

    ```sh
    firebase deploy
    ```

    Firebase will provide you with a URL where your add-in is hosted, like `https://your-project-id.web.app`.

8. Update the URLs in your `manifest.xml` file to point to your add-in's new location on Firebase. Replace "https://localhost:3000" occurrences with your Firebase URL followed by the path to your files.

9. Side-load your add-in in Outlook by following the instructions in the [Outlook add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
