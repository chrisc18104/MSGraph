using Azure.Identity;
using Microsoft.Graph;

const String filePath = @"C:\Temp\"; 
const String theSubject = "MSGraph Example Incoming";
const string sender_match = @"sender@domain.tld";

var settings = Settings.LoadSettings();

ClientSecretCredential credentials = new ClientSecretCredential(settings.TenantId /*TenantId*/, 
                                             settings.ClientId /*ClientID*/,
                                             settings.ClientSecret /*Client Secret*/);
GraphServiceClient graphClient = new GraphServiceClient(
                             credentials, 
                             settings.GraphUserScopes);

// the Active Directory user reference.
var theUser = graphClient.Me;
//var theUser = graphClient.Users[@"user@domain.tld"];
// to get data about the referenced user, make a request.
User theUserData = await theUser
	.Request()
	.GetAsync();
Console.WriteLine($"Hello, {theUserData?.DisplayName}!");
// the mailbox reference for the referenced user.
var theMailbox = theUser.MailFolders["Inbox"];
// use the messages collection reference to Request your list of messages.
IMailFolderMessagesCollectionPage theMail = await theMailbox.Messages
	.Request()
	.GetAsync();
foreach (var msg in theMail)
{
    Console.WriteLine( msg.HasAttachments + " " + msg.Subject.Contains(theSubject) + " " + msg.Subject + " " + msg.From.EmailAddress.Address.ToString());
	// filter for only the messages you want. 
	// note: to get the text email address, look in EmailAddress.Address
	if (msg.Subject.Contains(theSubject) && msg.HasAttachments == true && msg.From.EmailAddress.Address.ToString() == sender_match)
	{
		// Use the message's attachment reference to get the list of attachments.
		IMessageAttachmentsCollectionPage? attach = await theMailbox.Messages[msg.Id].Attachments
			.Request()
			.GetAsync();
		foreach(Attachment? theFile in attach)
		{
			Console.WriteLine( " " + theFile.Name + " " + theFile.ODataType + " " + theFile.IsInline);
			// filter for the type of attachments you wish. 
			if (theFile is FileAttachment && theFile.IsInline == false && theFile.Name.EndsWith(@".xlsx"))
			{
				// process attachment data through datastream into file on your local drive. [or shared drive. you do you.]
				FileAttachment fileAttachment = theFile as FileAttachment;
				byte[] stream = fileAttachment.ContentBytes;
				Stream stream1 = new MemoryStream(stream);
				FileStream fileStream = System.IO.File.Create(filePath + fileAttachment.Name);
				stream1.Seek(0, SeekOrigin.Begin);
				stream1.CopyTo(fileStream);
				fileStream.Close();

				/*
				 * Magick happens here! Your magick!
				 */
				
				// Send file with attachment starts here.
				// process file through datastream into data to go with attachment.
				String fileName = filePath + @"returning document.png";
				FileStream fs = new FileStream(fileName, FileMode.Open);
				Byte[] bytes = new Byte[fs.Length];
				fs.Seek(0, SeekOrigin.Begin);
				fs.Read(bytes, 0, bytes.Length);
				//String b64 = Convert.ToBase64String(bytes);

				// a fresh message to send.
				Message theResult = new Message
				{
					Subject = "MSGraph Example Outgoing on " + DateTime.Now.ToString(@"yyyy-MM-dd"),
					Body = new ItemBody
					{
						ContentType = BodyType.Html,
						Content = @"<p>This is example outgoing new message done through MSGraph.</p>"
					},
					ToRecipients = new List<Recipient>()
					{
						new Recipient
						{
							EmailAddress = new EmailAddress
							{
								Address = "to-ser@domain.tld"
							}
						}
					},
					CcRecipients = new List<Recipient>()
					{
						new Recipient
						{
							EmailAddress = new EmailAddress
							{
								Address = "ccer@domain.tld"
							}
						}, // more than one address is allowed in any list!
						new Recipient
						{
							EmailAddress = new EmailAddress
							{
								Address = "ccer2@domain.tld"
							}
						}
					},
					/*
					BccRecipients = new List<Recipient>()
					{
						new Recipient
						{
							EmailAddress = new EmailAddress
							{
								Address = "bccer@domain.tld"
							}
						}
					},
					*/
				};
				 
				// create attachment and include processed data.
				FileAttachment attachment = new FileAttachment
				{
					Name = @"Windows.png",
					ContentType = @"image/png",
					ContentBytes = bytes
				};

				// initialize attachment collection and add your attachment.
				theResult.Attachments = new MessageAttachmentsCollectionPage();	
				theResult.Attachments.Add(attachment);
				
				var saveToSentItems = true;

				// send a new message
				await theUser
					.SendMail(theResult,saveToSentItems)
					.Request()
					.PostAsync();
				
				// recipients of forwarded message.
                List<Recipient> toRecipients = new List<Recipient>()
                {
	                new Recipient
	                {
		                EmailAddress = new EmailAddress
		                {
			                Address = "fwder@domain.tld"
		                }
	                }
                };
                    
				String comment = "This is example outgoing forwarded message done through MSGraph.";
				// attach to existing message and forward to toRecipients 
                await theMailbox.Messages[msg.Id].Attachments
	                .Request()
	                .AddAsync(attachment);
                await theMailbox.Messages[msg.Id]
	                .Forward(toRecipients,null,comment)
	                .Request()
	                .PostAsync();

				// create a draft reply-all message. CreateReply replies only to the original sender.
                Message draftReplyAll = await theMailbox.Messages[msg.Id]
	                .CreateReplyAll(null,@"This is example outgoing reply message done through MSGraph.")
	                .Request()
	                .PostAsync();
				//attach file to created draft
				//Note: replies and replies-all lose existing attachments and cannot
				//have attachments added without creating draft.

				// Add CC and BCC recpipients to draft via update.
				List<Recipient> BccRecipients = new List<Recipient>()
				{
					new Recipient
					{
						EmailAddress = new EmailAddress
						{
							Address = "kevinr@allbrandonline.com"
						}
					},
					new Recipient
					{
						EmailAddress = new EmailAddress
						{
							Address = "stevef@allbrandonline.com"
						}
					},
					new Recipient
					{
						EmailAddress = new EmailAddress
						{
							Address = "croper@allbrandonline.com"
						}
					}
				};
				await theMailbox.Messages[draftReplyAll.Id]
					.Request()
					.UpdateAsync(new Message {BccRecipients = BccRecipients});
				// add attachment to draft
                await theMailbox.Messages[draftReplyAll.Id].Attachments
	                .Request()
	                .AddAsync(attachment);				
				// send the created draft
				await theMailbox.Messages[draftReplyAll.Id]
					.Send()
					.Request()
					.PostAsync();
				/* 
				attempting to send a created draft via SendMail barks about 'odata.context'
				await theUser
					.SendMail(draftReplyAll,saveToSentItems)
					.Request()
					.PostAsync();
				*/
				// Send file with attachment ends here.
			}                 
		}
		// mark item as read
        await theMailbox.Messages[msg.Id]
	        .Request()
	        .UpdateAsync(new Message {IsRead = true});
		//move item to deleted.
        var destinationId = "deleteditems";
        await theMailbox.Messages[msg.Id]
	        .Move(destinationId)
	        .Request()
	        .PostAsync();

        /* -- I assume this is a 'hard delete' instead of 'move to deleted items', since this method is invoked for multiple objects.
        await theMailbox.Messages[msg.Id]
	        .Request()
	        .DeleteAsync() ;
        */
	}
}
