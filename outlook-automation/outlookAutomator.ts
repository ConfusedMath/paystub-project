import { Client, LargeFileUploadTask, LargeFileUploadTaskOptions, FileUpload } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import fs from "fs-extra";
import path from "path";
import { config } from "dotenv";

config();

// Define the interface for the employee configurations
export interface EmployeeDraftConfig {
    name: string;
    email: string;
    filename: string;
    bufferData: Buffer;
}

// Define the structure of the execution report
export interface DraftExecutionReport {
    email: string;
    status: "success" | "error";
    draftId?: string;
    error?: string;
}

// Cache directory setup
const CACHE_DIR = path.join(__dirname, ".tmp", "email-cache");

/**
 * Initializes the Microsoft Graph client.
 * NOTE: Replace this with proper OAuth2 Token Credential Provider as per your exact auth flow.
 */
function getGraphClient(accessToken?: string): Client {
    return Client.init({
        authProvider: (done: (error: any, accessToken: string | null) => void) => {
            // If you have a specific token acquisition logic (e.g., Client Credentials),
            // put it here. For now, we assume an access token is passed or available via env.
            const token = accessToken || process.env.GRAPH_ACCESS_TOKEN;
            if (!token) {
                return done(new Error("No access token provided"), null);
            }
            done(null, token);
        }
    });
}

/**
 * Safely purges a file from the cache directory.
 */
async function purgeCache(filePath: string): Promise<void> {
    try {
        if (await fs.pathExists(filePath)) {
            await fs.remove(filePath);
            console.log(`[Cache] Purged cached file: ${filePath}`);
        }
    } catch (error) {
        console.error(`[Cache] Error purging file ${filePath}:`, error);
    }
}

/**
 * Main batch function to create Outlook drafts with attachments.
 * 
 * @param employees Array of employee configurations.
 * @param accessToken Optional OAuth2 access token if not using env vars.
 * @returns Array of DraftExecutionReports.
 */
export async function create_outlook_drafts(
    employees: EmployeeDraftConfig[],
    accessToken?: string
): Promise<DraftExecutionReport[]> {

    const client = getGraphClient(accessToken);
    const reports: DraftExecutionReport[] = [];

    // Ensure cache directory exists
    await fs.ensureDir(CACHE_DIR);

    for (const emp of employees) {
        const report: DraftExecutionReport = { email: emp.email, status: "error" };
        const cachedFilePath = path.join(CACHE_DIR, `${Date.now()}_${emp.filename}`);

        try {
            // 1. Cache the buffer data locally as a safety mechanism
            await fs.writeFile(cachedFilePath, emp.bufferData);
            console.log(`[Cache] Saved ${emp.filename} to ${cachedFilePath}`);

            // 2. Read file size
            const stats = await fs.stat(cachedFilePath);
            const fileSize = stats.size;
            const threshold = 4 * 1024 * 1024; // 4MB

            // 3. Create the empty draft message first
            const draftPayload = {
                subject: `Paystub for ${emp.name}`,
                body: {
                    contentType: "HTML",
                    content: `Hello ${emp.name},<br><br>Please find your paystub attached.`
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: emp.email
                        }
                    }
                ]
            };

            const createdDraft = await client.api("/me/messages").post(draftPayload);
            const draftId = createdDraft.id;

            // 4. Attach the file
            if (fileSize < threshold) {
                // Base64 direct upload for files < 4MB
                const base64Data = emp.bufferData.toString("base64");
                const attachmentPayload = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    name: emp.filename,
                    contentBytes: base64Data
                };

                await client.api(`/me/messages/${draftId}/attachments`).post(attachmentPayload);
                console.log(`[Attachment] Uploaded ${emp.filename} directly (Base64) to draft ${draftId}`);
            } else {
                // Large file upload session for files >= 4MB
                const uploadSessionPayload = {
                    AttachmentItem: {
                        attachmentType: "file",
                        name: emp.filename,
                        size: fileSize
                    }
                };

                const uploadSession = await client
                    .api(`/me/messages/${draftId}/attachments/createUploadSession`)
                    .post(uploadSessionPayload);

                // Use the FileUpload interface provided by graph client
                const fileObject = new FileUpload(emp.bufferData, emp.filename, fileSize);

                const options: LargeFileUploadTaskOptions = {
                    rangeSize: 320 * 1024 // 320 KB chunk size
                };

                const uploadTask = new LargeFileUploadTask(client, fileObject, uploadSession, options);
                await uploadTask.upload();

                console.log(`[Attachment] Uploaded ${emp.filename} via chunked session to draft ${draftId}`);
            }

            // 5. Success cleanup
            await purgeCache(cachedFilePath);

            report.status = "success";
            report.draftId = draftId;

        } catch (error: any) {
            console.error(`[Error] Failed processing draft for ${emp.email}:`, error);
            report.error = error.message || String(error);
        } finally {
            reports.push(report);
        }
    }

    return reports;
}
