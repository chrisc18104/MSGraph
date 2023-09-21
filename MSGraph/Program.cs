using All_Brand_Common;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Me;
using Microsoft.Graph.Me.MailFolders.Item;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item;

const String filePath = @"C:\Temp\"; 
const String theSubject = "MSGraph Example Incoming";
const string sender_match = @"sender@domain.tld";

Settings? settings = Settings.LoadSettings();

ClientSecretCredential credentials = new ClientSecretCredential(settings.TenantId /*TenantId*/, 
                                             settings.ClientId /*ClientID*/,
                                             settings.ClientSecret /*Client Secret*/);
GraphServiceClient graphClient = new GraphServiceClient(
                             credentials, 
                             settings.GraphUserScopes);

// the Active Directory user reference.
MeRequestBuilder theUser = graphClient.Me;
//UserItemRequestBuilder theUser = graphClient.Users[@"user@domain.tld"];
// to get data about the referenced user, make a request.
User? theUserData = await theUser
	.GetAsync();
Console.WriteLine($"Hello, {theUserData?.DisplayName}!");
// the mailbox reference for the referenced user.
MailFolderItemRequestBuilder theMailbox = theUser.MailFolders["Inbox"];
// use the messages collection reference to Request your list of messages.
MessageCollectionResponse? theMail = await theMailbox.Messages
	.GetAsync();
foreach (Message? msg in theMail.Value.FindAll(msg => msg.Subject.Contains(theSubject) && msg.HasAttachments == true && msg.From.EmailAddress.Address.ToString() == sender_match))
{
    Console.WriteLine( msg.HasAttachments + " " + msg.Subject.Contains(theSubject) + " " + msg.Subject + " " + msg.From.EmailAddress.Address.ToString());
	// filter for only the messages you want. 
	// note: to get the text email address, look in EmailAddress.Address
		// Use the message's attachment reference to get the list of attachments.

		AttachmentCollectionResponse? attach = await theMailbox.Messages[msg.Id].Attachments
			.GetAsync();
		foreach(Attachment? theFile in attach.Value.FindAll(theFile => theFile is FileAttachment && theFile.IsInline == false && theFile.Name.EndsWith(@".xlsx")))
		{
			Console.WriteLine( " " + theFile.Name + " " + theFile.OdataType + " " + theFile.IsInline);
			// filter for the type of attachments you wish. 
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
				};
				 
				// create attachment and include processed data.
				FileAttachment attachment = new FileAttachment
				{
					Name = @"Windows.png",
					ContentType = @"image/png",
					ContentBytes = bytes
				};

				// initialize attachment collection and add your attachment.
				List<Attachment> files = new List<Attachment>();
				files.Add(attachment);
				theResult.Attachments = files;
				
				bool saveToSentItems = true;

				var body = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
				{
					Message = theResult,
					SaveToSentItems = saveToSentItems
				};

				await graphClient.Users[ABLogin.abaEmail]
					.SendMail
					.PostAsync(body);

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
	                .PostAsync(attachment);
				var theForwardPostRequestBody = new Microsoft.Graph.Me.MailFolders.Item.Messages.Item.Forward.ForwardPostRequestBody
				{
					ToRecipients = toRecipients,
					Comment = comment
				};
               await theMailbox.Messages[msg.Id]
	                .Forward
	                .PostAsync(theForwardPostRequestBody);

				// create a draft forward message. 
				var theCreateForwardPostRequestBody = new Microsoft.Graph.Me.MailFolders.Item.Messages.Item.CreateForward.CreateForwardPostRequestBody
				{
					ToRecipients = toRecipients,
					Comment = comment
				};
                Message? draftForwardAll = await theMailbox.Messages[msg.Id]
	                .CreateForward
	                .PostAsync(theCreateForwardPostRequestBody);

				// create a draft reply-all message. CreateReply replies only to the original sender.
 				var theCreateReplyAllPostRequestBody = new Microsoft.Graph.Me.MailFolders.Item.Messages.Item.CreateReplyAll.CreateReplyAllPostRequestBody
				{
					Comment = comment
				};
                Message? draftReplyAll = await theMailbox.Messages[msg.Id]
	                .CreateReplyAll
	                .PostAsync(theCreateReplyAllPostRequestBody);

				//Section below pertains to CreateForward, CreateReply, and CreateReplyAll messages.
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
							Address = "name1@domain.tld"
						}
					},
					new Recipient
					{
						EmailAddress = new EmailAddress
						{
							Address = "name2@domain.tld"
						}
					},
					new Recipient
					{
						EmailAddress = new EmailAddress
						{
							Address = "name3@domain.tld"
						}
					}
				};
				await theMailbox.Messages[draftReplyAll.Id]
					.PatchAsync(new Message {BccRecipients = BccRecipients});
				// add attachment to draft
                await theMailbox.Messages[draftReplyAll.Id].Attachments
	                .PostAsync(attachment);				
				// send the created draft
				await theMailbox.Messages[draftReplyAll.Id]
					.Send
					.PostAsync();
				// Send file with attachment ends here.
		}
		// mark item as read
        await theMailbox.Messages[msg.Id]
	        .PatchAsync(new Message {IsRead = true});
		//move item to deleted.
        string destinationId = "deleteditems";
		var theMovePostRequestBody = new Microsoft.Graph.Me.MailFolders.Item.Messages.Item.Move.MovePostRequestBody
		{
			DestinationId = destinationId
		};
        await theMailbox.Messages[msg.Id]
	        .Move
	        .PostAsync(theMovePostRequestBody);

        /* -- I assume this is a 'hard delete' instead of 'move to deleted items', since this method is invoked for multiple objects.
        await theMailbox.Messages[msg.Id]
	        .DeleteAsync() ;
        */
}
