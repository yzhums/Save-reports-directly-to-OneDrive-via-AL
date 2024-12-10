pageextension 50115 PostedSalesInvoicesExt extends "Posted Sales Invoices"
{
    actions
    {
        addafter("&Invoice")
        {
            action(GenerateReportsAndUploadToOneDrive)
            {
                ApplicationArea = All;
                Caption = 'Generate Reports And Upload To OneDrive';
                Image = Import;
                Promoted = true;
                PromotedCategory = Process;
                trigger OnAction()
                var
                    OneDriveHandler: Codeunit OneDriveHandler;
                    SalesInvHeader: Record "Sales Invoice Header";
                    i: Integer;
                begin
                    i := 0;
                    SalesInvHeader.Reset();
                    CurrPage.SetSelectionFilter(SalesInvHeader);
                    if SalesInvHeader.FindSet() then
                        repeat
                            OneDriveHandler.UploadFilesToOneDrive(SalesInvHeader);
                            i += 1;
                        until SalesInvHeader.Next() = 0;
                    if i > 0 then
                        Message('%1 files uploaded to OneDrive successfully.', i);
                end;
            }
        }
    }
}
codeunit 50120 OneDriveHandler
{
    procedure UploadFilesToOneDrive(SalesInvHeader: Record "Sales Invoice Header")
    var
        HttpClient: HttpClient;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        Headers: HttpHeaders;
        ContentHeader: HttpHeaders;
        RequestContent: HttpContent;
        JsonResponse: JsonObject;
        AuthToken: SecretText;
        OneDriveFileUrl: Text;
        ResponseText: Text;
        FileContent: InStream;
        TempBlob: Codeunit "Temp Blob";
        FileName: Text;
        MimeType: Text;
        SalesInvHeader2: Record "Sales Invoice Header";
        ReportSelection: Record "Report Selections";
        TempReportSelections: Record "Report Selections" temporary;
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();
        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        // Generate the report and save it as a PDF file
        SalesInvHeader2.Get(SalesInvHeader."No.");
        SalesInvHeader2.SetRecFilter();
        ReportSelection.FindReportUsageForCust(Enum::"Report Selection Usage"::"S.Invoice", SalesInvHeader2."Bill-to Customer No.", TempReportSelections);
        Clear(TempBlob);
        TempReportSelections.SaveReportAsPDFInTempBlob(TempBlob, TempReportSelections."Report ID", SalesInvHeader2, TempReportSelections."Custom Report Layout Code", Enum::"Report Selection Usage"::"S.Invoice");
        TempBlob.CreateInStream(FileContent);
        FileName := Format(SalesInvHeader2."No." + '.pdf');
        MimeType := 'application/pdf';

        // Define the OneDrive folder URL
        // delegated permissions
        //OneDriveFileUrl := 'https://graph.microsoft.com/v1.0/me/drive/root/children';
        // application permissions (replace with the actual user principal name)
        OneDriveFileUrl := 'https://graph.microsoft.com/v1.0/users/Admin@2qcj3x.onmicrosoft.com/drive/root:/OneDriveAPITest/' + FileName + ':/content';
        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(OneDriveFileUrl);
        HttpRequestMessage.Method := 'PUT';
        HttpRequestMessage.GetHeaders(Headers);
        //Headers.Remove('Authorization');
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));
        RequestContent.GetHeaders(ContentHeader);
        ContentHeader.Clear();
        ContentHeader.Add('Content-Type', MimeType);
        HttpRequestMessage.Content.WriteFrom(FileContent);
        // Send the HTTP request
        if HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then begin
            // Log the status code for debugging
            //Message('HTTP Status Code: %1', HttpResponseMessage.HttpStatusCode());
            if HttpResponseMessage.IsSuccessStatusCode() then begin
                //HttpResponseMessage.Content.ReadAs(ResponseText);
                //JsonResponse.ReadFrom(ResponseText);
                //Message(ResponseText);
            end else begin
                //Report errors!
                HttpResponseMessage.Content.ReadAs(ResponseText);
                Error('Failed to upload files to OneDrive: %1 %2', HttpResponseMessage.HttpStatusCode(), ResponseText);
            end;
        end else
            Error('Failed to send HTTP request to OneDrive');
    end;

    procedure GetOAuthToken() AuthToken: SecretText
    var
        ClientID: Text;
        ClientSecret: Text;
        TenantID: Text;
        AccessTokenURL: Text;
        OAuth2: Codeunit OAuth2;
        Scopes: List of [Text];
    begin
        ClientID := 'b4fe1687-f1ab-4bfa-b494-0e2236ed50bd';
        ClientSecret := 'huL8Q~edsQZ4pwyxka3f7.WUkoKNcPuqlOXv0bww';
        TenantID := '7e47da45-7f7d-448a-bd3d-1f4aa2ec8f62';
        AccessTokenURL := 'https://login.microsoftonline.com/' + TenantID + '/oauth2/v2.0/token';
        Scopes.Add('https://graph.microsoft.com/.default');
        if not OAuth2.AcquireTokenWithClientCredentials(ClientID, ClientSecret, AccessTokenURL, '', Scopes, AuthToken) then
            Error('Failed to get access token from response\%1', GetLastErrorText());
    end;
}
