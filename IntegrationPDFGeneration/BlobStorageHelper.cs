using System;
using System.Collections.Generic;

using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using static IntegrationPDFGeneration.CommonDefinitions;

namespace IntegrationPDFGeneration
{
    class BlobStorageHelper
    {
        private string connectionString = "";
        private string containerRef = "";
        public BlobStorageHelper(String connectionString, String containerRef)
        {
            this.connectionString = connectionString;
            this.containerRef = containerRef;
        }
        public String uploadDocument(byte[] fileContent, String uploadFileName, String folderName)
        {
            CloudStorageAccount storageacc = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient blobClient = storageacc.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference(containerRef);

            CloudBlobDirectory directory = container.GetDirectoryReference(folderName);

            CloudBlockBlob blockBlob = directory.GetBlockBlobReference(uploadFileName);
            blockBlob.Metadata["keepUntil"] = getKeepUntilDate();

            blockBlob.UploadFromByteArray(fileContent, 0, fileContent.Length);

            return blockBlob.Uri.AbsoluteUri;
        }


        public String getFolderLocation (String folderName)
        {
            CloudStorageAccount storageacc = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient blobClient = storageacc.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference(containerRef);
            CloudBlobDirectory directory = container.GetDirectoryReference(folderName);

            return directory.Uri.AbsoluteUri;

        }

        public String getKeepUntilDate()
        {
            return DateTime.Now.AddMinutes(10).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        }

        public List<multiDocument> uploadMultipleDocuments(List<PDFStream> pdfStreams, String folderName)
        {
            CloudStorageAccount storageacc = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient blobClient = storageacc.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference(containerRef);

            container.CreateIfNotExists();

            CloudBlobDirectory directory = container.GetDirectoryReference(folderName);
            List<multiDocument> documentUrls = new List<multiDocument>();

            foreach (PDFStream pdfs in pdfStreams)
            {
                CloudBlockBlob blockBlob = directory.GetBlockBlobReference(pdfs.fileName);

                blockBlob.Metadata["keepUntil"] = getKeepUntilDate();

                blockBlob.UploadFromByteArray(pdfs.stream, 0, pdfs.stream.Length);
                //documentUrls.Add(blockBlob.Uri.AbsoluteUri);

                documentUrls.Add(new multiDocument(blockBlob.Uri.AbsoluteUri, pdfs.identifier));

            }

            return documentUrls;
        }


        public Boolean checkConnection()
        {
            CloudStorageAccount storageacc = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient blobClient = storageacc.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference(containerRef);

            return container.Exists();
        }
    }
}
