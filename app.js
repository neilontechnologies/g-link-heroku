
const express = require('express');
const cors = require('cors');
const app = express();

const { google } = require('googleapis');
const { Readable } = require('stream');

const { OAuth2 } = google.auth;
const SCOPE = ['https://www.googleapis.com/auth/drive'];

app.use(express.urlencoded({extended:true}));
app.use(express.json());
app.use(express.text());
app.use(cors());
const bodyParser = require('body-parser');
app.use(bodyParser.json())

const crypto = require('crypto');

const XMLHttpRequest = require('xmlhttprequest').XMLHttpRequest;

// Use to authenticate heroku access key
app.use((req, res, next) => {
  const contentType = req.headers['content-type'];
  let salesforceAuthenticationInfo;
  const apiKey = process.env.API_KEY;
  if(contentType == 'text/plain'){
    const decryptedPayload = decryptAES256(req.body, apiKey.substring(0, 32));
    salesforceAuthenticationInfo = JSON.parse(decryptedPayload);
  } else{
	  salesforceAuthenticationInfo = req.body;
  }
  const { heroku_api_key } = salesforceAuthenticationInfo;

  if(heroku_api_key === apiKey){
    next(); 
  } else{
    res.status(403).send('Forbidden: Invalid Heroku API Key');
  }
});

// This service is used to upload salesforce files and attachments into Google Drive
app.post('/uploadsalesforcefile', async (req, res) => {
  try{
    const contentType = req.headers['content-type'];
    let salesforceAuthenticationInfo;
    if(contentType == 'text/plain'){
	  const apiKey = process.env.API_KEY;
      const decryptedPayload = decryptAES256(req.body, apiKey.substring(0, 32));
      salesforceAuthenticationInfo = JSON.parse(decryptedPayload);
    } else{
	    salesforceAuthenticationInfo = req.body;
    }
	
    // Get all headers from apex
	const {
	  google_drive_client_id,
	  google_drive_secret_id,
	  sf_client_id,
	  sf_client_secret,
	  sf_username,
	  sf_password,
	  google_drive_file_title,
	  sf_parent_id,
	  google_drive_folder_key,
	  google_drive_bucket_name,
	  sf_content_document_id,
	  sf_file_size,
	  sf_file_id,
	  sf_content_document_link_id,
	  sf_namespace,
	  sf_delete_file,
	  sf_create_log,
	  g_file,
	  google_drive_file_meta_data,
	  google_drive_refresh_token,
	  google_drive_folder_id,
	  sf_instance_url,
	  sf_token,
      sf_bulk_job_id,
      google_drive_token,
      storage,
      sharepoint_upload_url
	} = salesforceAuthenticationInfo;

    // We are sending the request immediately because we cannot wait untill the whole migration is completed. It will timeout the API request in Apex.
    res.send(`Heroku service to migrate Salesforce File has been started successfully.`);

    // Get salesforce response
    const migrateSalesforceResult = migrateSalesforce(sf_file_id, google_drive_client_id, google_drive_secret_id, google_drive_refresh_token, sf_client_id, sf_client_secret, sf_username, sf_password, google_drive_bucket_name, google_drive_folder_key, google_drive_file_title, sf_file_size, 
    sf_content_document_id, sf_parent_id, sf_content_document_link_id, sf_namespace, sf_delete_file, sf_create_log, g_file, google_drive_file_meta_data, google_drive_folder_id, sf_instance_url, sf_token, sf_bulk_job_id, google_drive_token, storage, sharepoint_upload_url);
  } catch(error){
    console.log(error);
  }
});

// This methiod is used to handle all combine methods
const migrateSalesforce = async (sfFileId, googleDriveAccessKey, googleDriveSecretKey, googleDriveRefreshToken, sfClientId, sfClientSecret, sfUsername, sfPassword, googleDriveBucketName, googleDriveFolderKey, googleDriveFileTitle, sfFileSize, sfContentDocumentId, sfParentId, sfContentDocumentLinkId, sfNamespace, sfDeleteFile, sfCreateLog, gFile, googleDriveFileMetadata, googleDriveFolderId, sfInstanceUrl, sfToken, sfBulkJobId, googleDriveToken, storage, sharepointUploadUrl) =>{
  let salesforceAccessToken;
  let instanceUrl;

  // Check token
  if(sfToken == null){
    // Get access token of salesforce
    const salesforceTokenResponse = await getSalesforceToken(sfClientId, sfClientSecret, sfUsername, sfPassword, sfInstanceUrl);

    // Check if access token and instance URL are available or not
    if(!salesforceTokenResponse.accessToken || !salesforceTokenResponse.instanceUrl){
      return;
    } else {
      salesforceAccessToken = salesforceTokenResponse.accessToken;
      instanceUrl = salesforceTokenResponse.instanceUrl
    }
  } else{
    salesforceAccessToken = sfToken;
    instanceUrl = sfInstanceUrl;
  }


  // Get access token authetication with google drive
  let googleDriveAccessToken;
  if(storage === 'SharePoint'){
    googleDriveAccessToken = googleDriveToken;
  } else if(storage === 'Google Drive'){
    // Get access token authetication with google drive
    googleDriveAccessToken = await createOAuthClient(googleDriveAccessKey, googleDriveSecretKey, googleDriveRefreshToken);

    // Check google access token is null or not
    if(googleDriveAccessToken == null){
      const failureReason = 'Google Drive authentication failed';
      const createFileMigrationLogResult = createLogs(salesforceAccessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
    }
  }

  // Check required parameters
  if(((sfFileSize &&  sfFileId) || sfBulkJobId) && (googleDriveFolderKey || sfParentId) && googleDriveFileTitle){
    let getSalesforceFileResult;
    if(sfBulkJobId != null){
        // Get salesforce data information 
        getSalesforceFileResult = await getSalesforceData(salesforceAccessToken, instanceUrl, sfBulkJobId, sfNamespace, sfCreateLog);
    } else {
        // Get salesforce file information 
        getSalesforceFileResult = await getSalesforceFile(salesforceAccessToken, instanceUrl, sfFileId, sfContentDocumentLinkId, sfNamespace, sfCreateLog);
    }
    console.log(sfParentId);
	if(sfParentId != null){
      // Check if google drive folder id is available for parentId or not
      const { getRecordHomeFolderResult } = await getRecordHomeFolder(salesforceAccessToken, instanceUrl, sfParentId, sfFileId, sfContentDocumentLinkId, sfNamespace, sfCreateLog, sfBulkJobId);
	  
      // Check reponse
      if(getRecordHomeFolderResult.sObjects != null && getRecordHomeFolderResult.sObjects.length > 0){
		// Set google folder path
		googleDriveFolderKey = getRecordHomeFolderResult.sObjects[0][sfNamespace + 'Google_Folder_Path__c'];
		
		// Set googlde folder id
	    googleDriveFolderId = getRecordHomeFolderResult.sObjects[0][sfNamespace + 'Google_Drive_Folder_Id__c'];
		
		// Set bucket name 
		googleDriveBucketName = getRecordHomeFolderResult.sObjects[0][sfNamespace + 'Bucket_Name__c'];
      } else{
          // Prepare failure rason with error message of API
          const failureReason = 'Your request to create G-Folder for the record failed. ERROR: ' + getRecordHomeFolderResult.message ;

        if(sfCreateLog){
            // Create File Migration Logs
            const createFileMigrationLogResult = await createLogs(salesforceAccessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
        }
      }
    }
	console.log(storage);
	console.log('googleDriveFolderKey-->' + googleDriveFolderKey);
	if(googleDriveFolderKey != null){
      // Prepare google drive folder path
      const googleDriveFolderPath = googleDriveBucketName + '/' + googleDriveFolderKey;

      // Create google drive file path
      const googleDriveFilePath = googleDriveFolderKey + '/' + googleDriveFileTitle

      // Check if storage is Google Drive and google folder id not available
      if(storage === 'Google Drive' && googleDriveFolderId == null){
		// Create google drive folder using google drive folder path  
        let createGoogleDriveFolderResult = await createGoogleDriveFolder(salesforceAccessToken, instanceUrl, googleDriveFolderPath, sfFileId, sfContentDocumentLinkId, sfNamespace, sfCreateLog, sfBulkJobId);
		
		//  Check response
		if(createGoogleDriveFolderResult != null && createGoogleDriveFolderResult.code == 200 && createGoogleDriveFolderResult.data != null){
            // Get google drive folder id
			googleDriveFolderId = createGoogleDriveFolderResult.data.split('/').pop();
        } else{
          // Prepare failure rason with error message of API
          const failureReason = 'Your request to create Google Drive Folders failed. ERROR: ' + createGoogleDriveFolderResult.message;
		  
          if(sfCreateLog){
            // Create File Migration Logs 
            const createFileMigrationLogResult = createLogs(salesforceAccessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
          }
        }
	  }
	  
	  console.log(googleDriveFolderId);
      // Check folder is created or not
      if(storage === 'SharePoint' || googleDriveFolderId != null){
		console.log('Upload');
        let response;
        if(storage === 'SharePoint'){
            // Upload file into SharePoint
            response = await uploadFileToSharePoint(googleDriveAccessToken, sharepointUploadUrl, gFile, getSalesforceFileResult, salesforceAccessToken, instanceUrl, sfFileId, sfContentDocumentLinkId, sfCreateLog, sfNamespace, sfBulkJobId);
        } else{
            // Upload file into google drive , 
            response = await uploadFileToGoogleDrive(googleDriveAccessToken, getSalesforceFileResult, googleDriveFolderId, googleDriveFileTitle, gFile, sfNamespace, salesforceAccessToken, instanceUrl, sfFileId, sfContentDocumentLinkId, sfCreateLog, googleDriveFileMetadata, sfBulkJobId);
        }
        
        // Check response
		if(response.status == 200){
          if(response && response.data && response.data.id){
            // Get google drive file id
            const googleDriveFileId = response.data.id;
            const googleDriveFileSize = response.data.size;

            // Create G-file record if file is successfully uploaded into google drive
            const createGFilesInSalesforceResult = await createGFilesInSalesforce(salesforceAccessToken, instanceUrl, googleDriveBucketName, googleDriveFilePath, googleDriveFileSize, sfContentDocumentId, sfFileId, sfContentDocumentLinkId, sfNamespace, sfDeleteFile, sfCreateLog, gFile, googleDriveFileId, sfBulkJobId);
          }
        }
      }
    }
  } else{
    if(sfCreateLog){
      // Prepare failure rason with error message of API
      let failureReason = '';

      if(sfBulkJobId && sfBulkJobId.trim() !== ''){
        failureReason = 'Salesforce Bulk API Job Id, Google Drive Bucket Name, or Google Drive Folder Path is missing.';
      } else{
        failureReason = 'Salesforce File Id, Salesforce File Size, Google Drive Bucket Name, or Google Drive Folder Path is missing.';
      }

      // Create File Migration Logs
      const createFileMigrationLogResult = await createLogs(salesforceAccessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
      throw new Error(failureReason);
    }
  }
}

// This method is used to get access token of Salesforce org and instance url of the org
const getSalesforceToken = (sfClientId, sfClientSecret, sfUsername, sfPassword, sfInstanceUrl) => {
  return new Promise((resolve, reject) => {
    const postData = `grant_type=password&client_id=${sfClientId}&client_secret=${sfClientSecret}&username=${sfUsername}&password=${sfPassword}`;
    const xhr = new XMLHttpRequest();

    xhr.open('POST', sfInstanceUrl + '/services/oauth2/token', true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');

    xhr.onreadystatechange = function(){
      if(xhr.readyState === 4){
        const response = JSON.parse(xhr.responseText);
        if(xhr.status === 200){
          resolve({
            accessToken: response.access_token,
            instanceUrl: response.instance_url
          });
        } else {
          reject(new Error('We are not able to get the Salesforce Authentication Token. This happens if the Salesforce Client Id, Client Secret, User Name, Password or Security Token is invalid. ERROR: ' + response.error_description));
        }
      }
    };

    xhr.onerror = function(e){
      reject(new Error('We are not able to get the Salesforce Authentication Token. This happens if the Salesforce Client Id, Client Secret, User Name, Password or Security Token is invalid. ERROR: ' + e.message));
    };

    xhr.send(postData);
  });
};

// This method is used to get salesforce file information with the help of access token of that org, URL, provided salesforce file id  
const getSalesforceFile = async (accessToken, instanceUrl, sfFileId, sfContentDocumentLinkId, sfNamespace, sfCreateLog) => {
  var url;
  // Prepare url of attachments or content document
  if(sfFileId.startsWith('00P')){
    url = `${instanceUrl}/services/data/v58.0/sobjects/Attachment/${sfFileId}/Body`;
  } else {
    url = `${instanceUrl}/services/data/v58.0/sobjects/ContentVersion/${sfFileId}/VersionData`;
  }
  
  // To authenticate salesforce
  try {
    const response = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    // Returns the response status code
    if(!response.ok){
      throw new Error(`We are not able to fetch the Salesforce File Content. ERROR: ${response.statusText}`);
    } else{
	  const chunks = [];
	  for await (const chunk of response.body) {
	    chunks.push(chunk);
	  }
	  
	  if(chunks.length > 0){
	    const buffer = Buffer.concat(chunks);
	    return buffer;
	  } else{
	    throw new Error('Salesforce File body is empty.');
	  }
    }
  } catch(error){
    // Create File Migration Logs
    if(sfCreateLog){
      const createFileMigrationLogResult = await createFileMigrationLog(accessToken, instanceUrl, null, sfFileId, sfContentDocumentLinkId, error.message, sfNamespace);
	}
    console.error(error);
    throw error;
  }
};

// This method is used to fetch Salesforce data with the help of bulk api job id
const getSalesforceData = async (accessToken, instanceUrl, sfBulkJobId, sfNamespace, sfCreateLog) => {
  try {
	  // Prepare url for Salesforce data export
    const url = `${instanceUrl}/services/data/v60.0/jobs/query/${sfBulkJobId}/results`;
	
	  // To authenticate salesforce
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });
	
	  // Returns the response status code
    if (!response.ok) {
      // Handle error response
      throw new Error(`We are not able to fetch the Salesforce Data. ERROR: ${response.statusText}`);
    } else {
      // Handle successful response
      const chunks = [];
      for await (const chunk of response.body) {
        chunks.push(chunk);
      }

      if(chunks.length > 0){
        const buffer = Buffer.concat(chunks);
        return buffer;
	    } else{
		    throw new Error('Salesforce Data body is empty.');
	    }
    }

  } catch (error) {
	// Create File Migration Logs
    if(sfCreateLog){
      const createFileMigrationLogResult = await createLogs(accessToken, instanceUrl, sfBulkJobId, null, null, error.message, sfNamespace);
	}
    console.error(error);
    throw error;
  }
};

// This method used to create record home folder for parent id
const getRecordHomeFolder = async (accessToken, instanceUrl, sfParentId, sfFileId, sfContentDocumentLinkId, sfNamespace, sfCreateLog, sfBulkJobId) => {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    let url;

    // Check namespace is available or not
    if(sfNamespace){
      url = `${instanceUrl}/services/apexrest/NEILON2/GLink/v1/recordfolder/${sfParentId}`;
    } else{
      url = `${instanceUrl}/services/apexrest/GLink/v1/recordfolder/${sfParentId}`;
    }

    xhr.open('GET', url, true); 
    xhr.setRequestHeader('Authorization', `Bearer ${accessToken}`);
    xhr.setRequestHeader('Content-Type', 'application/json');  

    xhr.onload = function() {
      if(xhr.readyState === 4){
        const response = JSON.parse(xhr.responseText);
        console.log(response);
        if(xhr.status === 200){
          resolve({
            getRecordHomeFolderResult: response
          });  // Resolve the Promise on success
        } else {
          // Prepare error messsage
          const failureReason = 'Your request to create G-Folder for the record failed. ERROR: ' + response[0].message;

          // Check sf create log is true or false
          if (sfCreateLog) {
			      // Ensure the log is created before rejecting the promise
            createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
          }
        }
      }
    };

    xhr.onerror = function(e) {
      // Prepare failure rason with error message of API
      const failureReason = 'Your request to create G-Folder for the record failed. ERROR: ' + e;

      // Check sf create log is true or false
      if(sfCreateLog){
		    // Create File Migration Logs
        const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
      }

      // Handle network error
      reject(new Error(failureReason));
    };
    xhr.send();
  });
};


// This method used to create G-Files record in salesforce
const createGoogleDriveFolder = async (accessToken, instanceUrl, googleDriveFolderPath, sfFileId, sfContentDocumentLinkId, sfNamespace, sfCreateLog, sfBulkJobId) => {
  return new Promise((resolve, reject) => {
    let url;
    const xhr = new XMLHttpRequest();

    //Check namespace is available or not
    if(sfNamespace != ''){
      url = `${instanceUrl}/services/apexrest/NEILON2/GLink/v1/creategoogledrivefolders/`;
    } else {
      url = `${instanceUrl}/services/apexrest/GLink/v1/creategoogledrivefolders/`;
    }
    
    // Prepare body
    var textBody = googleDriveFolderPath;

    // Open the request
    xhr.open('POST', url, true);
    xhr.setRequestHeader('Authorization', `Bearer ${accessToken}`);
    xhr.setRequestHeader('Content-Type', 'text/plain');
    
    xhr.onload = function() {
      if (xhr.readyState === 4) {
        const response = JSON.parse(xhr.responseText);
        console.log(response);
        if (xhr.status === 200) {
          resolve({
            createGoogleDriveFolderResult: response
          });  // Resolve the Promise on success
        } else {
          // Prepare failure rason with error message of API
          const failureReason = 'Your request to create Google Drive Folder failed. ERROR: ' + response[0].message;

          if(sfCreateLog){
			      // Create File Migration Logs
            const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
          }
        }
      }
    };
    
    // Send the request with the JSON body
    xhr.send(textBody);
  });
};

// This method used to create G-Files record in salesforce
const createGFilesInSalesforce = async (accessToken, instanceUrl, googleDriveBucketName, googleDriveFilePath, sfFileSize, sfContentDocumentId, sfFileId, sfContentDocumentLinkId, sfNamespace, sfDeleteFile, sfCreateLog, gFile, googleDriveFileId, sfBulkJobId) => {
  return new Promise((resolve, reject) => {
    let url;
    const xhr = new XMLHttpRequest();

    // Check namespace is available or not
    if(sfNamespace !== ''){
      url = `${instanceUrl}/services/apexrest/NEILON2/GLink/v1/creategfiles/`;
    } else {
      url = `${instanceUrl}/services/apexrest/GLink/v1/creategfiles/`;
    }
    
    var body = [];

    // Check g-file is availbe or not
    if(!gFile){
      gFile = {};
    }

    gFile[sfNamespace + 'Bucket_Name__c'] = googleDriveBucketName;
    gFile[sfNamespace + 'Google_File_Path__c'] = googleDriveFilePath;
    gFile[sfNamespace + 'Size__c'] = sfFileSize;
    gFile[sfNamespace + 'Content_Document_Id__c'] = sfContentDocumentId;
    gFile[sfNamespace + 'Export_Attachment_Id__c'] = sfFileId;
    gFile[sfNamespace + 'Google_Drive_File_Id__c'] = googleDriveFileId;
    body.push(gFile);

    // Open the request
    xhr.open('POST', url, true);
    xhr.setRequestHeader('Authorization', `Bearer ${accessToken}`);
    xhr.setRequestHeader('Content-Type', 'application/json');

    if(sfDeleteFile){
      xhr.setRequestHeader('delete-salesforce-file', 'true');
    }

    // Handle the response
    xhr.onload = function() {
      if (xhr.readyState === 4) {  
        const response = JSON.parse(xhr.responseText);
        console.log(response);
        if (xhr.status === 200) {
          if(response.sObjects && response.sObjects.length > 0 && !response.sObjects[0].Id){
            // Prepare failure rason with error message of API
            const failureReason = 'Your request to create G-Files in Salesforce failed. ERROR: ' + response.sObjects[0][sfNamespace + 'Description__c'];

            // Check sf create log is true or false
            if(sfCreateLog){
              // Create File Migration Logs
              const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
            }
          } else{
            resolve(response);
          }
        } else {
          // Prepare failure rason with error message of API
          const failureReason = 'Your request to create G-Files in Salesforce failed. ERROR: ' + response[0].message;
          
          // Check sf create log is true or false
          if(sfCreateLog){
			      // Create File Migration Logs
            const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
          }

          reject(new Error(failureReason));
        }
      }
    };

    // Handle network errors
    xhr.onerror = function(e) {
      // Prepare failure rason with error message of API
      const failureReason = 'Your request to create G-Files in Salesforce failed. ERROR: ' + e;

      // Check sf create log is true or false
      if(sfCreateLog){
		    // Create File Migration Logs
        const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
      }
      reject(new Error(failureReason));
    };

    // Send the request with the JSON body
    xhr.send(JSON.stringify(body));
  });
};

// This method used to create Salesforce File Migration Log record in salesforce 
const createLogs = (accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace) => {
  return new Promise((resolve, reject) => {
    let url;
    const xhr = new XMLHttpRequest();
    let body = {};
	
    // Check if this is data export job
    const isBulkJob = sfBulkJobId && sfBulkJobId.trim() !== '';
	
    // Decide endpoint
    if(sfNamespace != ''){
	    if(isBulkJob){
		    url = `${instanceUrl}/services/apexrest/NEILON2/GLink/v1/setdataexportjobstatus/`;
	    } else{
		    url = `${instanceUrl}/services/apexrest/NEILON2/GLink/v1/createmigrationlog/`;
	    }
    } else {
	    if(isBulkJob){
		    url = `${instanceUrl}/services/apexrest/GLink/v1/setdataexportjobstatus/`;
	    } else{
		    url = `${instanceUrl}/services/apexrest/GLink/v1/createmigrationlog/`;
	    }
    }

    xhr.open('POST', url, true);
    xhr.setRequestHeader('Authorization', `Bearer ${accessToken}`);
    xhr.setRequestHeader('Content-Type', 'application/json');

    // Prepare body
    if(isBulkJob){
      // Bulk Job case
      body = {
        JobId: sfBulkJobId,
        Status: 'Failed',
        Message: failureReason
      };
    } else{
      // Check file type is attachment or content document link
      if (sfFileId.startsWith('00P')) {
        body.SalesforceFileId = sfFileId;
      } else {
        body.SalesforceFileId = sfContentDocumentLinkId;
      }
      body.FailureReason = failureReason;
    }

    // Handle the response
    xhr.onload = function() {
      if (xhr.readyState === 4) {  
        const response = JSON.parse(xhr.responseText);
        console.log(response);
        if (xhr.status === 200) {
          resolve(response);
        } else {
          return;
        }
      }
    };

    // Handle network errors
    xhr.onerror = function(e) {
      return;
    };

    // Send the request with the JSON body
    xhr.send(JSON.stringify(body));
  });
};

// This function is used to create authentication with google drive
async function createOAuthClient(clientId, clientSecret, refreshToken) {
  const oauth2Client = new OAuth2(clientId, clientSecret);
  oauth2Client.setCredentials({ refresh_token: refreshToken });
  return oauth2Client;
}

// This function will upload the desired file to google drive folder
async function uploadFileToGoogleDrive(authClient, buffer, googleDriveFolderId, googleDriveFileTitle, gFile, sfNamespace, accessToken, instanceUrl, sfFileId, sfContentDocumentLinkId, sfCreateLog, googleDriveFileMetadata, sfBulkJobId) {
  return new Promise((resolve, reject) => {
    // Authenticate with google
    const drive = google.drive({ version: 'v3', auth: authClient });
 
    // Get meta tags
    var fileMetaTags = {};
    const metatype = 'google';

    // Create google drive file metadata
    if(googleDriveFileMetadata){
      // Prepare google drive metadata map
      Object.entries(googleDriveFileMetadata).forEach(([filedAPIName, value]) => {
        var fieldAPI = filedAPIName;
        var metaFieldAPI = 'x-' + metatype + '-meta-' + fieldAPI.toLowerCase();
        if (googleDriveFileMetadata[fieldAPI] !== undefined && googleDriveFileMetadata[fieldAPI] !== null) {
          fileMetaTags[metaFieldAPI] = googleDriveFileMetadata[fieldAPI].toString();
        } else {
            fileMetaTags[metaFieldAPI] = '';
        }
      })
    }

    // Prepare metadata to store in google drive file
    const googleDriveFolderIds = [];
    googleDriveFolderIds.push(googleDriveFolderId);

    const fileMetaData = {
      name: googleDriveFileTitle,
      parents: googleDriveFolderIds, 
      mimeType: gFile[sfNamespace + 'Content_Type__c'],
      properties: fileMetaTags
    };
    
    // Create a readable stream from the buffer
    const bufferStream = Readable.from(buffer);

    // Prepare media for google drive file
    const media = {
      body: bufferStream,
      mimeType: gFile[sfNamespace + 'Content_Type__c'],
    };

    // Method to upload file in google drive
    drive.files.create(
      {
        resource: fileMetaData,
        media,
        fields: 'id, size',
      },
      async (error, file) => {
        if(error){
          // Check sf create log is true or false
          if(sfCreateLog){
            // Prepare error message
            const failureReason = 'Your request to upload file in Google Drive has failed. ' + error;

		    // Create File Migration Logs
            const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
          }
          return;
        }

        if(gFile[sfNamespace + 'Public_On_Google__c']){
          try {
            await drive.permissions.create({
              fileId: file.data.id,
              requestBody: {
                role: 'reader',
                type: 'anyone',
              },
            });
          } catch (permissionError) {
            const failureReason = 'Your request to make file public on Google Drive failed. ERROR: ' + permissionError;

            // Check sf create log is true or false
            if (sfCreateLog) {
			  // Create File Migration Logs
              const createFileMigrationLogResult = createLogs(accessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, failureReason, sfNamespace);
            }
            return;
          }
        }
        resolve(file);
      }
    );
  });
}

// This function will upload the desired file to share point folder
async function uploadFileToSharePoint(sharePointToken, sharepointUploadInfo, gFile, buffer, salesforceAccessToken, instanceUrl, sfFileId, sfContentDocumentLinkId, sfCreateLog, sfNamespace, sfBulkJobId){
    try {
	  console.log('Endpoint-->' + sharepointUploadInfo.url);
      if(sharepointUploadInfo.type === 'UploadSessionUrl'){
        // Upload large file using multipart upload
		const uploadUrl = sharepointUploadInfo.url;
        const chunkSize = 10 * 1024;

        let start = 0;
        let uploadResponse;

        while(start < buffer.length){
          const end = Math.min(start + chunkSize, buffer.length);
          const chunk = buffer.slice(start, end);

          uploadResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
              'Content-Length': chunk.length,
              'Content-Range':
                `bytes ${start}-${end - 1}/${buffer.length}`
            },
            body: chunk
          });

          if(!uploadResponse.ok){
            let errorData = await uploadResponse.text();
            throw new Error(errorData);
          }
          start = end;
        }

        let responseUploadData = await uploadResponse.json();
        return {
          status: 200,
          data: responseUploadData
        };
      } else if(sharepointUploadInfo.type === 'Endpoint'){
		console.log('sharePointToken-->' + sharePointToken);
		console.log('Content_Type__c-->' + gFile[sfNamespace + 'Content_Type__c']);

        // Upload small file using single part upload
        let uploadResponse = await fetch(sharepointUploadInfo.url, {
          method: 'PUT',
          headers: {
            'Authorization': `Bearer ${sharePointToken}`,
            'Content-Type': gFile[sfNamespace + 'Content_Type__c']
          },
          body: buffer
        });
		
		if(!uploadResponse.ok){
          let errorData = await uploadResponse.text();
		  console.log(errorData);
          throw new Error(errorData);
        }
		  
        let responseData = await uploadResponse.json();
        return {
          status: 200,
          data: responseData
        };
      }
    } catch(error){
        if(sfCreateLog){
          // Create File Migration Logs
          const createFileMigrationLogResult =  createLogs(salesforceAccessToken, instanceUrl, sfBulkJobId, sfFileId, sfContentDocumentLinkId, error.message, sfNamespace);
        }
    }
}

// This function will be used to decrypt the payload
function decryptAES256(encryptedBase64, keyString) {
  // Convert the key and encrypted text to buffers
  const key = Buffer.from(keyString, 'utf8'); // 32 bytes for AES256
  const encryptedData = Buffer.from(encryptedBase64, 'base64');

  // Salesforce prepends IV (16 bytes) to the ciphertext
  const iv = encryptedData.subarray(0, 16);
  const ciphertext = encryptedData.subarray(16);

  // Create decipher
  const decipher = crypto.createDecipheriv('aes-256-cbc', key, iv);
  let decrypted = decipher.update(ciphertext);
  decrypted = Buffer.concat([decrypted, decipher.final()]);

  // Convert back to string
  return decrypted.toString('utf8');
}

// This service is used to upload salesforce files and attachments into Google Drive from local host
app.get('/', async (req, res) => {
  // Example usage
  try {
    // Replace these values with your own Salesforce Connected App credentials
    const sfFileId = '{SALESFORCE_CONTENT_VERSION_ID}'; 
    const googleDriveClientId = '{GOOGLE_DRIVE_CLIENT_ID}';
    const googleDriveClientSecretId = '{GOOGLE_DRIVE_CLIENT_SECRET_ID}';
    const sfClientId = '{SALESFORCE_CLIENT_ID}';
    const sfClientSecret = '{SALESFORCE_CLIENT_SECRET_KEY}';
    const sfUsername = '{SALESFORCE_USERNAME}';
    const sfPassword = '{SALESFORCE_PASSWORD}';
    const googleDriveBucketName = '{GOOGLE_DRIVE_BUCKET_NAME}';
    const sfFileSize = '{SALESFORCE_FILE_SIZE}';
    const sfContentDocumentId = '{SALESFORCE_CONTENT_DOCUMENT_ID}';
    const googleDriveFolderKey = '{GOOGLE_DRIVE_FOLDER_KEY}'
    const googleDriveFileTitle = 'GOOGLE_DRIVE_FILE_TITLE';
    const sfParentId = '{SALESFORCE_PARENT_ID}';
    const sfContentDocumentLinkId = '{SALESFORCE_CONTENT_DOCUMENT_LINK_ID}';
    const sfNamespace = '{SALESFORCE_NAMESPACE}';
    const sfDeleteFile = '{SALESFORCE_DELETE_FILE}';
    const sfCreateLog = '{SALESFORCE_CREATE_LOG}';
    const gFile = '{G_FILE}';
    const googleDriveFileMetadata = '{GOOGLE_DRIVE_FILE_METADATA}';
    const googleDriveRefreshToken = '{GOOGLE_DRIVE_REFRESH_TOKEN}';
    const googleDriveFolderId = '{GOOGLE_DRIVE_FOLDER_ID}';
	const sfInstanceUrl = '{SALESFORCE_INSTANCE_URL}';
    const sfToken = '{SALESFORCE_TOKEN}';
	const sfBulkJobId = '{SALESFORCE_BULK_JOB_ID}';
    const googleDriveToken = "{GOOGLE_DRIVE_TOKEN}";
    const storage = '{STORAGE}';
    const sharePointUploadInfo = '{SHAREPOINT_UPLOAD_INFO}';

    // We are sending the request immediately because we cannot wait untill the whole migration is completed. It will timeout the API request in Apex.
    res.send(`Heroku service to migrate Salesforce File has been started successfully.`);
    
    const reponse = await migrateSalesforce (sfFileId, googleDriveClientId, googleDriveClientSecretId, googleDriveRefreshToken, sfClientId, sfClientSecret, sfUsername, sfPassword, googleDriveBucketName, googleDriveFolderKey, googleDriveFileTitle, sfFileSize, sfContentDocumentId, sfParentId, sfContentDocumentLinkId, sfNamespace, sfDeleteFile, sfCreateLog, gFile, googleDriveFileMetadata, googleDriveFolderId, sfInstanceUrl, sfToken, sfBulkJobId, googleDriveToken, storage, sharePointUploadInfo);
  } catch (error) {
    console.error(error);
  }
});

const port = process.env.PORT || 3008;
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
