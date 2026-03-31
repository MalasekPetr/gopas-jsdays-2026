import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getGraphClient } from "./graphClient";
import { TARGET_USER_EMAIL, AUDIT_LIST_NAME } from "./config";

interface GrantAccessBody {
  siteId: string;
  listId: string;
  itemId: string;
}

async function grantAccess(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  try {
    const { siteId, listId, itemId } = (await req.json()) as GrantAccessBody;
    const client = getGraphClient();

    // Step 1a: Get the default drive ID for the site
    console.log(`Getting default drive for site ${siteId}...`);
    const drive = await client.api(`/sites/${siteId}/drive`).get();
    const driveId = drive.id;

    // Step 1b: Grant read permission on the item via sharing invite
    console.log(`Granting read access to ${TARGET_USER_EMAIL} on item ${itemId}...`);
    await client.api(`/drives/${driveId}/items/${itemId}/invite`).post({
      recipients: [{ email: TARGET_USER_EMAIL }],
      roles: ["read"],
      requireSignIn: true,
      sendInvitation: false,
    });

    // Step 2: Create audit log entry in the SharePoint list
    console.log(`Writing audit log to list ${AUDIT_LIST_NAME}...`);
    await client.api(`/sites/${siteId}/lists/${AUDIT_LIST_NAME}/items`).post({
      fields: {
        Title: itemId,
        GrantedTo: TARGET_USER_EMAIL,
        Timestamp: new Date().toISOString(),
      },
    });

    return {
      status: 200,
      jsonBody: { success: true, itemId, grantedTo: TARGET_USER_EMAIL },
    };
  } catch (err: any) {
    return {
      status: 500,
      jsonBody: { error: err.message },
    };
  }
}

app.http("grantAccess", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "grant-access",
  handler: grantAccess,
});
