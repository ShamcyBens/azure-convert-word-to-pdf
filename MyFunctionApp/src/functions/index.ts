import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { WordsApi, ConvertDocumentRequest } from "asposewordscloud";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function for converting Word document to PDF.');

    if (req.body && req.body.fileUrl) {
        const fileUrl = req.body.fileUrl;

        try {
            // Initialize Aspose.Words API client
            const wordsApi = new WordsApi("your-aspose-client-id", "your-aspose-client-secret");

            // Prepare the request for document conversion
            const request = new ConvertDocumentRequest({
                document: fileUrl,
                format: "pdf"
            });

            // Convert the document to PDF
            const pdfFile = await wordsApi.convertDocument(request);

            // Return the PDF file as a response
            context.res = {
                status: 200,
                headers: {
                    'Content-Type': 'application/pdf',
                    'Content-Disposition': `attachment; filename=${req.body.fileName.replace('.docx', '.pdf')}`
                },
                body: pdfFile.body
            };
        } catch (error) {
            context.res = {
                status: 500,
                body: `Error converting document: ${error.message}`
            };
        }
    } else {
        context.res = {
            status: 400,
            body: "Please provide a valid fileUrl in the request body."
        };
    }
};

export default httpTrigger;