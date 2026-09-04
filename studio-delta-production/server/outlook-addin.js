const ADDIN_ID = "c8e4d2a1-7b93-4f0e-9c12-1a2b3c4d5e6f";

function publicOrigin(req) {
  const proto = String(req.headers["x-forwarded-proto"] || req.protocol || "https").split(",")[0].trim() || "https";
  const host = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
  if (!host) return "https://localhost";
  return proto + "://" + host;
}

function manifestXml(origin) {
  const icon = origin + "/outlook-addin/icon";
  const taskpane = origin + "/outlook-addin/taskpane.html";
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="MailApp">
  <Id>${ADDIN_ID}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Studio Delta</ProviderName>
  <DefaultLocale>en-ZA</DefaultLocale>
  <DisplayName DefaultValue="Studio Delta"/>
  <Description DefaultValue="Save this Outlook email onto a Studio Delta enquiry. The email stays in Outlook; the app only keeps a link."/>
  <IconUrl DefaultValue="${icon}-32.png"/>
  <HighResolutionIconUrl DefaultValue="${icon}-80.png"/>
  <SupportUrl DefaultValue="${origin}/outlook-addin"/>
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
    <AppDomain>https://appsforoffice.microsoft.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.5"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${taskpane}"/>
        <RequestedHeight>420</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
  <DisableEntityHighlighting>true</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.5">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="taskpaneUrl"/>
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="sdGroup">
                <Label resid="groupLabel"/>
                <Control xsi:type="Button" id="sdSaveButton">
                  <Label resid="buttonLabel"/>
                  <Supertip>
                    <Title resid="buttonLabel"/>
                    <Description resid="buttonDesc"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon16"/>
                    <bt:Image size="32" resid="icon32"/>
                    <bt:Image size="80" resid="icon80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="taskpaneUrl"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon16" DefaultValue="${icon}-16.png"/>
        <bt:Image id="icon32" DefaultValue="${icon}-32.png"/>
        <bt:Image id="icon80" DefaultValue="${icon}-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="taskpaneUrl" DefaultValue="${taskpane}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Studio Delta"/>
        <bt:String id="buttonLabel" DefaultValue="Studio Delta"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="buttonDesc" DefaultValue="Save this email onto a Studio Delta enquiry. Nothing is downloaded."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

function addinHeaders(res) {
  res.setHeader("Cache-Control", "no-store, max-age=0");
  res.setHeader(
    "Content-Security-Policy",
    "frame-ancestors 'self' https://*.office.com https://*.office365.com https://outlook.office.com https://outlook.office365.com https://outlook.live.com"
  );
}

function mountOutlookAddin(app, publicDir) {
  const path = require("path");
  const dir = path.join(publicDir, "outlook-addin");
  app.get("/outlook-addin/manifest.xml", (req, res) => {
    addinHeaders(res);
    res.type("application/xml").send(manifestXml(publicOrigin(req)));
  });
  app.get(["/outlook-addin", "/outlook-addin/"], (_req, res) => {
    addinHeaders(res);
    res.sendFile(path.join(dir, "index.html"));
  });
  app.get("/outlook-addin/taskpane.html", (_req, res) => {
    addinHeaders(res);
    res.sendFile(path.join(dir, "taskpane.html"));
  });
  app.get("/outlook-addin/icon-:size.png", (req, res) => {
    const size = String(req.params.size || "");
    if (!/^(16|32|80)$/.test(size)) {
      res.status(404).end();
      return;
    }
    addinHeaders(res);
    res.type("image/png").sendFile(path.join(dir, "icon-" + size + ".png"));
  });
}

module.exports = { mountOutlookAddin, manifestXml, publicOrigin, ADDIN_ID };
