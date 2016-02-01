{***************************************************************************}
{ TIdSmtpEx component                                                       }
{ for Delphi base Indy9                                                     }
{                                                                           }
{ 继承TIdSMTP组件用于解决以下问题：                                         }
{   1、SendBody方法同时发送TIdText和TIdAttachment时TIdText发送失败问题      }
{   2、SendHeader主题长度被截断问题                                         }
{ written by mini188                                                        }
{ Email : mini188.com@qq.com                                                }
{                                                                           }
{***************************************************************************}

unit IdSMTPEx;

interface
uses
  SysUtils,
  IdComponent, IdSMTP, IdMessage, IdTCPConnection, IdTCPClient, IdMessageClient,
  IdBaseComponent;
type
  TIdSmtpEx = class(TIdSMTP)
  protected
    procedure SendHeader(AMsg: TIdMessage); override;
    procedure SendBody(AMsg: TIdMessage); override;
  end;


implementation
uses
  IdTCPStream,
  IdCoderHeader, IdCoderQuotedPrintable, IdGlobal, IdMessageCoderMIME,
  IdResourceStrings, IdHeaderList;


{ TIdSmtpEx }

/// <summary>
/// 发送内容
/// 此方法是为了解决发送TIdText 的bug
/// </summary>
/// <param name="AMsg"></param>
procedure TIdSmtpEx.SendBody(AMsg: TIdMessage);
var
  i: Integer;
  LAttachment: TIdAttachment;
  LBoundary: string;
  LDestStream: TIdTCPStream;
  LMIMEAttachments: boolean;
  ISOCharset: string;
  HeaderEncoding: Char;  { B | Q }
  TransferEncoding: TTransfer;

  procedure WriteTextPart(ATextPart: TIdText);
  var
    Data: string;
    i: Integer;
  begin
    if Length(ATextPart.ContentType) = 0 then
      ATextPart.ContentType := 'text/plain'; {do not localize}
    if Length(ATextPart.ContentTransfer) = 0 then
      ATextPart.ContentTransfer := 'quoted-printable'; {do not localize}
    WriteLn('Content-Type: ' + ATextPart.ContentType); {do not localize}
    WriteLn('Content-Transfer-Encoding: ' + ATextPart.ContentTransfer); {do not localize}
    WriteStrings(ATextPart.ExtraHeaders);
    WriteLn('');

    // TODO: Provide B64 encoding later
    // if AnsiSameText(ATextPart.ContentTransfer, 'base64') then begin
    //  LEncoder := TIdEncoder3to4.Create(nil);

    if AnsiSameText(ATextPart.ContentTransfer, 'quoted-printable') then
    begin
      for i := 0 to ATextPart.Body.Count - 1 do
      begin
        if Copy(ATextPart.Body[i], 1, 1) = '.' then
        begin
          ATextPart.Body[i] := '.' + ATextPart.Body[i];
        end;
        Data := TIdEncoderQuotedPrintable.EncodeString(ATextPart.Body[i] + EOL);
        if TransferEncoding = iso2022jp then
          Write(Encode2022JP(Data))
        else
          Write(Data);
      end;
    end

    else begin
      WriteStrings(ATextPart.Body);
    end;
    WriteLn('');
  end;

begin
  LMIMEAttachments := AMsg.Encoding = meMIME;
  LBoundary := '';

  InitializeISO(TransferEncoding, HeaderEncoding, ISOCharSet);
  BeginWork(wmWrite);
  try
    if AMsg.MessageParts.AttachmentCount > 0 then
    begin
      if LMIMEAttachments then
      begin
        WriteLn('This is a multi-part message in MIME format'); {do not localize}
        WriteLn('');
        if AMsg.MessageParts.RelatedPartCount > 0 then
        begin
          LBoundary := IndyMultiPartRelatedBoundary;
        end
        else begin
          LBoundary := IndyMIMEBoundary;
        end;
        WriteLn('--' + LBoundary);
      end
      else begin
        // It's UU, write the body
        WriteBodyText(AMsg);
        WriteLn('');
      end;

      if AMsg.MessageParts.TextPartCount >= 1 then//原先是'>'修改为'>='用于支持 TIdText的发送
      begin
        WriteLn('Content-Type: multipart/alternative; '); {do not localize}
        WriteLn('        boundary="' + IndyMultiPartAlternativeBoundary + '"'); {do not localize}
        WriteLn('');
        for i := 0 to AMsg.MessageParts.Count - 1 do
        begin
          if AMsg.MessageParts.Items[i] is TIdText then
          begin
            WriteLn('--' + IndyMultiPartAlternativeBoundary);
            DoStatus(hsStatusText,  [RSMsgClientEncodingText]);
            WriteTextPart(AMsg.MessageParts.Items[i] as TIdText);
            WriteLn('');
          end;
        end;
        WriteLn('--' + IndyMultiPartAlternativeBoundary + '--');
      end
      else begin
        if LMIMEAttachments then
        begin
          WriteLn('Content-Type: text/plain'); {do not localize}
          WriteLn('Content-Transfer-Encoding: 7bit'); {do not localize}
          WriteLn('');
          WriteBodyText(AMsg);
        end;
      end;

      // Send the attachments
      for i := 0 to AMsg.MessageParts.Count - 1 do
      begin
        if AMsg.MessageParts[i] is TIdAttachment then
        begin
          LAttachment := TIdAttachment(AMsg.MessageParts[i]);
          DoStatus(hsStatusText, [RSMsgClientEncodingAttachment]);
          if LMIMEAttachments then
          begin
            WriteLn('');
            WriteLn('--' + LBoundary);
            if Length(LAttachment.ContentTransfer) = 0 then
            begin
              LAttachment.ContentTransfer := 'base64'; {do not localize}
            end;
            if Length(LAttachment.ContentDisposition) = 0 then
            begin
              LAttachment.ContentDisposition := 'attachment'; {do not localize}
            end;
            if (LAttachment.ContentTransfer = 'base64') {do not localize}
              and (Length(LAttachment.ContentType) = 0) then
            begin
              LAttachment.ContentType := 'application/octet-stream'; {do not localize}
            end;
            WriteLn('Content-Type: ' + LAttachment.ContentType + ';'); {do not localize}
            WriteLn('        name="' + ExtractFileName(LAttachment.FileName) + '"'); {do not localize}
            WriteLn('Content-Transfer-Encoding: ' + LAttachment.ContentTransfer); {do not localize}
            WriteLn('Content-Disposition: ' + LAttachment.ContentDisposition +';'); {do not localize}
            WriteLn('        filename="' + ExtractFileName(LAttachment.FileName) + '"'); {do not localize}
            WriteStrings(LAttachment.ExtraHeaders);
            WriteLn('');
          end;
          LDestStream := TIdTCPStream.Create(Self);
          try
            TIdAttachment(AMsg.MessageParts[i]).Encode(LDestStream);
          finally
            FreeAndNil(LDestStream);
          end;
          WriteLn('');
        end;
      end;
      if LMIMEAttachments then
      begin
        WriteLn('--' + LBoundary + '--');
      end;
    end
    // S.G. 21/2/2003: If the user added a single texpart message without filling the body
    // S.G. 21/2/2003: we still need to send that out
    else
    if (AMsg.MessageParts.TextPartCount > 1) or
       ((AMsg.MessageParts.TextPartCount = 1) and (AMsg.Body.Count = 0)) then
    begin
      WriteLn('This is a multi-part message in MIME format'); {do not localize}
      WriteLn('');
      for i := 0 to AMsg.MessageParts.Count - 1 do
      begin
        if AMsg.MessageParts.Items[i] is TIdText then
        begin
          WriteLn('--' + IndyMIMEBoundary);
          DoStatus(hsStatusText, [RSMsgClientEncodingText]);
          WriteTextPart(AMsg.MessageParts.Items[i] as TIdText);
        end;
      end;
      WriteLn('--' + IndyMIMEBoundary + '--');
    end

    else begin
      DoStatus(hsStatusText, [RSMsgClientEncodingText]);
      // Write out Body
      //TODO: Why just iso2022jp? Why not someting generic for all MBCS? Or is iso2022jp special?
      if TransferEncoding = iso2022jp then
      begin
        for i := 0 to AMsg.Body.Count - 1 do
        begin
          if Copy(AMsg.Body[i], 1, 1) = '.' then
          begin
            WriteLn('.' + Encode2022JP(AMsg.Body[i]));
          end

          else begin
            WriteLn(Encode2022JP(AMsg.Body[i]));
          end;
        end;
      end

      else begin
        WriteBodyText(AMsg);
      end;
    end;
  finally
    EndWork(wmWrite);
  end;
end;

procedure TIdSmtpEx.SendHeader(AMsg: TIdMessage);
var
  LHeaders: TIdHeaderList;
begin
  LHeaders := AMsg.GenerateHeader;
  try
    //解决标题过长时导致的收件方解码错误问题
    LHeaders.Text := StringReplace(LHeaders.Text, #13#10#13#10, #13#10, [rfReplaceAll]);
    WriteStrings(LHeaders);
  finally
    FreeAndNil(LHeaders);
  end;
end;

end.
