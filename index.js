// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <ProgramSnippet>
const readline = require('readline-sync');

const settings = require('./appSettings');
const graphHelper = require('./graphHelper');

// <InitializeGraphSnippet>
function initializeGraph(settings) {
  graphHelper.initializeGraphForUserAuth(settings, (info) => {
    // Display the device code message to
    // the user. This tells them
    // where to go to sign in and provides the
    // code to use.
    console.log(info.message);
  });
}
// </InitializeGraphSnippet>


async function listFilesAsync() {
  try {
    const response = await graphHelper.listFilesAsync();
    return await response;
  } catch (error) {
    // console.log(error)
    console.error('Error listing files:', error.message);
    throw error;
  }
}


async function getUsersWithAccessAsync(fileId) {
  try {
    const response = await graphHelper.getUsersWithAccessAsync(fileId);
    return response;
  } catch (error) {
    console.error('Error getting users with access:', error.message);
    throw error;
  }
}

const express = require('express');
const fs = require('fs');
const fetch = require('node-fetch');

const app = express();
const port = 3000;

initializeGraph(settings);


app.get('/', async (req, res) => {
  try {
    // const accessToken = await getAccessToken();
    const files = await listFilesAsync();


    // Render HTML using a template engine like ejs or pug for cleaner code
    res.send(`<!DOCTYPE html>
    <html>
    <head>
      <title>OneDrive File Manager</title>
      <style>          
        #selected-file-info {
          width: 30%;
          position: fixed;
          right: 0;
          top: 0;
          height: 100%;
          background-color: #f0f0f0;
          padding: 10px;
          box-shadow: -2px 0 5px rgba(0, 0, 0, 0.1);
          overflow-y: auto;
        }
        table {
          width: 100%;
          border-collapse: collapse;
        }
        th, td {
          padding: 10px;
          border: 1px solid #ddd;
          text-align: left;
        }
        th {
          background-color: #f2f2f2;
        }
      </style>
    </head>
    <body>
      <h1>OneDrive Files</h1>
      <table>
        <thead>
          <tr>
            <th>File Name</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          <!-- Populate table with files -->
          ${files.value.map(file => `
            <tr>
              <td>${file.name}</td>
              <td>
                <button onclick="getFileUsers('${file.id}', '${file.name}')">Users with Access</button>
                ${file.folder ? '<button disabled>Download</button>' : `<button onclick="downloadFile('${file.id}')">Download</button>`}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div>Note: Currently folder download is not implemented, only files can be downloaded.</div>
      <div id="selected-file-info"></div>
    
      <script>
        let fileIdToUpdate = null; // Track the fileId for which we are fetching users
        let timerId = null; // Initialize timerId variable
    
        async function pollFileUsers(fileId, fileName) {
          try {
            const interval = 3000; // Polling interval in milliseconds (e.g., every 30 seconds)
            timerId = setInterval(async () => {
              try {
                console.log('Polling users for file:', fileId);
                const response = await fetch(\`/users/\${fileId}\`);
                const data = await response.json(); // Parse the JSON response
    
                // Check if the fileId has changed or if there's new data
                if (fileId === fileIdToUpdate) {
                  console.log('Users fetched:', data);
    
                  // Update the UI with the fetched data
                  const selectedFileInfo = document.getElementById('selected-file-info');
                  selectedFileInfo.innerHTML = \`
                    <h2>Selected File: \${fileName}</h2>
                    <p>Users with access:</p>
                    <table>
                      <thead>
                        <tr>
                          <th>Username</th>
                          <th>Email</th>
                          <th>Role(s)</th>
                        </tr>
                      </thead>
                      <tbody>
                        \${data.map(permission => \`
                          <tr>
                            <td>\${permission.displayName}</td>
                            <td>\${permission.email}</td>
                            <td>\${permission.role}</td>
                          </tr>
                        \`).join('')}
                      </tbody>
                    </table>
                  \`;
                }
              } catch (error) {
                console.error('Error polling users:', error.message);
              }
            }, interval);
    
            return timerId;
          } catch (error) {
            console.error('Error starting polling:', error.message);
          }
        }
    
        async function getFileUsers(fileId, fileName) {
          try {
            if (timerId) {
              clearInterval(timerId); // Clear any existing polling
            }
    
            // Immediately fetch and display data
            const response = await fetch(\`/users/\${fileId}\`);
            const data = await response.json(); // Parse the JSON response
            console.log('Users fetched:', data);
    
            const selectedFileInfo = document.getElementById('selected-file-info');
            selectedFileInfo.innerHTML = \`
              <h2>Selected File: \${fileName}</h2>
              <p>Users with access:</p>
              <table>
                <thead>
                  <tr>
                    <th>Username</th>
                    <th>Email</th>
                    <th>Role(s)</th>
                  </tr>
                </thead>
                <tbody>
                  \${data.map(permission => \`
                    <tr>
                      <td>\${permission.displayName}</td>
                      <td>\${permission.email}</td>
                      <td>\${permission.role}</td>
                    </tr>
                  \`).join('')}
                </tbody>
              </table>
            \`;
    
            // Start polling for the new fileId
            fileIdToUpdate = fileId; // Set the fileId being polled
            timerId = await pollFileUsers(fileId, fileName);
          } catch (error) {
            console.error('Error fetching users:', error.message);
          }
        }
    
        async function downloadFile(fileId) {
          try {
            const response = await fetch(\`/download/\${fileId}\`);
            if (!response.ok) {
              throw new Error(\`HTTP error! Status: \${response.status}\`);
            }
    
            // Get filename from Content-Disposition header or use default
            const filename = getFilenameFromContentDisposition(response.headers);
    
            const blob = await response.blob(); // Get the response as Blob
    
            // Create a temporary URL to download the Blob
            const url = window.URL.createObjectURL(blob);
    
            // Create a link element to trigger the download
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = filename; // Set filename based on response or custom logic
            document.body.appendChild(a);
            a.click();
    
            // Cleanup: remove the temporary URL and link element
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
          } catch (error) {
            console.error('Error downloading file:', error.message);
          }
        }
    
        // Function to parse filename from Content-Disposition header
        function getFilenameFromContentDisposition(headers) {
          const contentDisposition = headers.get('content-disposition');
          let filename = 'downloaded_file';
          if (contentDisposition) {
            const match = contentDisposition.match(/filename="(.+)"/);
            if (match) {
              filename = match[1];
            }
          }
          return filename;
        }
      </script>
    </body>
    </html>    
    
    `);
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).send('Error retrieving files');
  }
});

app.get('/users/:fileId', async (req, res) => {
  try {
    const { fileId } = req.params;
    const {filePermissions, fileMetadata} = await getUsersWithAccessAsync(fileId);
    // Combine the owner and permissions into one list
    const usersWithAccess = filePermissions.value.map(permission => {
      const user = permission.grantedTo ? permission.grantedTo.user : null;
      return {
        displayName: user ? user.displayName : 'External user (Invited)',
        id: user ? user.id : null,
        email: permission.invitation ? permission.invitation.email : null,
        role: permission.roles.join(', ')
      };
    });

    // Add the owner to the list
    usersWithAccess.push({
      displayName: fileMetadata.createdBy.user.displayName,
      id: fileMetadata.createdBy.user.id,
      email: fileMetadata.createdBy.user.email,
      role: 'owner'
    });
    res.json(usersWithAccess);
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ error: 'Error retrieving users with access' });
  }
});

// Route to download a file by fileId
app.get('/download/:fileId', async (req, res) => {
  try {
    const { fileId } = req.params;

    // Download the file using graphHelper function
    const { metadata, content } = await graphHelper.downloadFileAsync(fileId);

    // Set appropriate headers for file download
    res.set({
      'Content-Type': metadata.file.mimeType,
      'Content-Disposition': `attachment; filename="${metadata.name}"`, // Set filename from metadata
    });

    // Send the file buffer as the response
    res.send(Buffer.from(content));
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ error: 'Error downloading file' });
  }
});


app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
