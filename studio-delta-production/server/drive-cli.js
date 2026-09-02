#!/usr/bin/env node
/**
 * Blocking Drive / Docs / Gmail helper. Called with spawnSync from the GAS shim.
 * Input JSON on stdin: { op, ... }
 * Output JSON on stdout.
 */
const { google } = require("googleapis");
const fs = require("fs");

function credentialsFromEnv() {
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  }
  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    return JSON.parse(fs.readFileSync(process.env.GOOGLE_APPLICATION_CREDENTIALS, "utf8"));
  }
  return null;
}

function getAuth(scopes) {
  const json = credentialsFromEnv();
  if (!json) throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON");
  return new google.auth.GoogleAuth({ credentials: json, scopes });
}

async function main() {
  const input = JSON.parse(fs.readFileSync(0, "utf8") || "{}");
  const scopes = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/gmail.send"
  ];
  const auth = getAuth(scopes);
  const drive = google.drive({ version: "v3", auth });
  const docs = google.docs({ version: "v1", auth });
  const driveOpts = { supportsAllDrives: true };
  const op = input.op;
  let out = { ok: true };

  if (op === "copyFile") {
    const copied = await drive.files.copy({
      fileId: input.fileId,
      requestBody: { name: input.name, parents: input.folderId ? [input.folderId] : undefined },
      fields: "id, webViewLink, name, mimeType, createdTime",
      ...driveOpts
    });
    out = { ok: true, id: copied.data.id, url: copied.data.webViewLink, name: copied.data.name, mimeType: copied.data.mimeType, createdTime: copied.data.createdTime };
  } else if (op === "replaceText") {
    const reqs = (input.replacements || []).map((r) => ({
      replaceAllText: { containsText: { text: r.tag, matchCase: true }, replaceText: String(r.value == null ? "" : r.value) }
    }));
    if (reqs.length) {
      await docs.documents.batchUpdate({ documentId: input.docId, requestBody: { requests: reqs } });
    }
  } else if (op === "exportPdf") {
    const pdf = await drive.files.export({ fileId: input.fileId, mimeType: "application/pdf" }, { responseType: "arraybuffer" });
    const buf = Buffer.from(pdf.data);
    out = { ok: true, pdfBase64: buf.toString("base64") };
  } else if (op === "uploadFile") {
    if (!input.path) throw new Error("uploadFile needs a path");
    const mime = input.mimeType || "application/octet-stream";
    const created = await drive.files.create({
      requestBody: {
        name: input.name,
        parents: input.folderId ? [input.folderId] : undefined,
        mimeType: mime
      },
      media: { mimeType: mime, body: fs.createReadStream(input.path) },
      fields: "id, webViewLink, name, mimeType, createdTime",
      ...driveOpts
    });
    out = { ok: true, id: created.data.id, url: created.data.webViewLink, name: created.data.name, mimeType: created.data.mimeType, createdTime: created.data.createdTime };
  } else if (op === "shareWithEmail") {
    await drive.permissions.create({
      fileId: input.fileId,
      requestBody: {
        type: "user",
        role: input.role || "writer",
        emailAddress: input.email
      },
      sendNotificationEmail: false,
      supportsAllDrives: true
    });
  } else if (op === "createFile") {
    const mime = input.mimeType || "text/plain";
    const body = input.contentBase64
      ? require("stream").Readable.from(Buffer.from(input.contentBase64, "base64"))
      : (input.content || "");
    const created = await drive.files.create({
      requestBody: {
        name: input.name,
        parents: input.folderId ? [input.folderId] : undefined,
        mimeType: mime
      },
      media: { mimeType: mime, body },
      fields: "id, webViewLink, name, mimeType, createdTime",
      ...driveOpts
    });
    out = { ok: true, id: created.data.id, url: created.data.webViewLink, name: created.data.name, mimeType: created.data.mimeType, createdTime: created.data.createdTime };
  } else if (op === "htmlToPdf") {
    const doc = await drive.files.create({
      requestBody: {
        name: input.name || "list",
        mimeType: "application/vnd.google-apps.document",
        parents: input.folderId ? [input.folderId] : undefined
      },
      media: { mimeType: "text/html", body: input.html || "<html></html>" },
      fields: "id",
      ...driveOpts
    });
    const pdf = await drive.files.export({ fileId: doc.data.id, mimeType: "application/pdf" }, { responseType: "arraybuffer" });
    const buf = Buffer.from(pdf.data);
    try {
      await drive.files.update({ fileId: doc.data.id, requestBody: { trashed: true }, ...driveOpts });
    } catch (ignore) {}
    out = { ok: true, pdfBase64: buf.toString("base64") };
  } else if (op === "listFiles") {
    const q = ["'" + input.folderId + "' in parents", "trashed = false"];
    if (input.mimeType) q.push("mimeType = '" + input.mimeType + "'");
    const files = [];
    let pageToken = undefined;
    do {
      const res = await drive.files.list({
        q: q.join(" and "),
        fields: "nextPageToken, files(id, name, mimeType, createdTime, webViewLink)",
        pageSize: 200,
        orderBy: "createdTime desc",
        pageToken,
        includeItemsFromAllDrives: true,
        supportsAllDrives: true
      });
      files.push.apply(files, res.data.files || []);
      pageToken = res.data.nextPageToken;
    } while (pageToken && files.length < 500);
    out = { ok: true, files: files.map((f) => ({ id: f.id, name: f.name, mimeType: f.mimeType, createdTime: f.createdTime, url: f.webViewLink, webViewLink: f.webViewLink })) };
  } else if (op === "getFileText") {
    const res = await drive.files.get({ fileId: input.fileId, alt: "media", supportsAllDrives: true }, { responseType: "text" });
    out = { ok: true, text: res.data };
  } else if (op === "trashFile") {
    await drive.files.update({ fileId: input.fileId, requestBody: { trashed: true }, supportsAllDrives: true });
  } else if (op === "renameFile") {
    await drive.files.update({ fileId: input.fileId, requestBody: { name: input.name }, supportsAllDrives: true });
  } else if (op === "listFoldersByName") {
    const res = await drive.files.list({
      q: "mimeType = 'application/vnd.google-apps.folder' and name = '" + String(input.name).replace(/'/g, "\\'") + "' and trashed = false",
      fields: "files(id, name, webViewLink)",
      pageSize: 10,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    out = { ok: true, files: res.data.files || [] };
  } else if (op === "createFolder") {
    const created = await drive.files.create({
      requestBody: { name: input.name, mimeType: "application/vnd.google-apps.folder" },
      fields: "id, name, webViewLink",
      supportsAllDrives: true
    });
    out = { ok: true, id: created.data.id, url: created.data.webViewLink, name: created.data.name };
  } else if (op === "sendMail") {
    const sender = process.env.GMAIL_SENDER;
    if (!sender) throw new Error("GMAIL_SENDER is not set");
    const boundary = "sd_" + Date.now();
    let raw = "From: " + sender + "\r\n";
    raw += "To: " + input.to + "\r\n";
    if (input.cc) raw += "Cc: " + input.cc + "\r\n";
    raw += "Subject: " + input.subject + "\r\n";
    raw += "MIME-Version: 1.0\r\n";
    if (input.attachments && input.attachments.length) {
      raw += "Content-Type: multipart/mixed; boundary=" + boundary + "\r\n\r\n";
      raw += "--" + boundary + "\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n" + (input.html || "") + "\r\n";
      for (const a of input.attachments) {
        raw += "--" + boundary + "\r\nContent-Type: " + (a.mime || "application/pdf") + "; name=\"" + a.name + "\"\r\n";
        raw += "Content-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename=\"" + a.name + "\"\r\n\r\n";
        raw += a.base64 + "\r\n";
      }
      raw += "--" + boundary + "--";
    } else {
      raw += "Content-Type: text/html; charset=UTF-8\r\n\r\n" + (input.html || "");
    }
    const encoded = Buffer.from(raw).toString("base64url");
    const json = credentialsFromEnv();
    const mailAuth = new google.auth.JWT({
      email: json.client_email,
      key: json.private_key,
      scopes: ["https://www.googleapis.com/auth/gmail.send"],
      subject: sender
    });
    const gmailAsUser = google.gmail({ version: "v1", auth: mailAuth });
    await gmailAsUser.users.messages.send({ userId: "me", requestBody: { raw: encoded } });
  } else {
    throw new Error("Unknown drive op " + op);
  }
  process.stdout.write(JSON.stringify(out));
}

main().catch((e) => {
  process.stdout.write(JSON.stringify({ ok: false, error: e.message || String(e) }));
  process.exit(0);
});
