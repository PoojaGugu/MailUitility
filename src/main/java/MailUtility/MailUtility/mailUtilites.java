package MailUtility.MailUtility;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.mail.Address;
import javax.mail.Flags;
import javax.mail.Flags.Flag;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMessage.RecipientType;
import javax.mail.internet.MimeMultipart;
import javax.mail.search.FlagTerm;


import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;

public class mailUtilites {
	private 	Session session;
	private Store store ;
	private Transport t;
	private Message message;
	private String EmailUser,EmailPassword,SubjectPattern;
	private Address[] fromAddress;
	
	public static void main(String args[]) throws MessagingException, IOException {
		if(args.length == 5)
		{
			mailUtilites receiver = new mailUtilites();
			receiver.mailSessionInit(args[0],args[1],args[4]);
			Multimap<String, String> mMap  =receiver.downloadEmailAttachments( args[2],args[3]);
			if(mMap!=null)
			{
				Set<String> keys = mMap.keySet();
				for (String key : keys)
				{


					Collection<String> files =  mMap.get(key);
					Iterator<String> iterator = files.iterator();
					while (iterator.hasNext()) 
					{
						System.out.println(iterator.next());
					}

				}
			}
		}
		else
		{
			System.out.println("Invalid arguments please specify in following order <emailid>,<password><InboxfolderName>,<DownloadDirectoryPath>,<Email Subject Pattern>");
			
		}
	}
	public void mailSessionInit(String emailID ,String emailpassword,String SubjectPattern) throws MessagingException, IOException
	{
		this.EmailUser=emailID;
		this.EmailPassword = emailpassword;
		this.SubjectPattern =  SubjectPattern;

		Properties props = new Properties();
		props.setProperty("mail.imap.ssl", "true");
		props.put("mail.smtp.host", "webmail.limitedbrands.com");
		props.put("mail.imaps.port", "993");
		props.put("mail.imaps.socketFactory.fallback", "false");
		props.put("mail.imaps.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		props.put("mail.store.protocol", "imaps");
		props.put("mail.smtp.host", "webmail.limitedbrands.com");
		props.put("mail.smtp.port", "587");
		props.put("mail.smtp.ssl.trust", "webmail.limitedbrands.com");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.connectiontimeout", "10000");
		this.session = Session.getDefaultInstance(props);
		/*  Session.getInstance(props, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(EmailUser, EmailPassword);
			}
		});*/
		this.store = session.getStore("imaps");
		System.out.println(store.isConnected());
		store.connect("webmail.limitedbrands.com",EmailUser,EmailPassword);
		System.out.println(store.isConnected());
		this.t = session.getTransport("smtp");
		this.message = new MimeMessage(session);
		createDirectory("Reports");


	}
	private static  String saveDirectory = "/home/content"; // directory to save the downloaded documents

	private static  String destDirectory ;
	/**
	 * Sets the directory where attached files will be stored.
	 * @param dir absolute path of the directory
	 */
	public void setSaveDirectory(String dir) {
		mailUtilites.saveDirectory = dir;
	}

	public void setdestDirectory(String dir) {
		mailUtilites.destDirectory = dir;
	}
	public String getdestDirectory() {
		return mailUtilites.destDirectory;
	}

	public  void sendMailWithAttachment(String fromAddress,String toAddress,String msgSubject,String msgBody,String[] attachFiles) throws UnsupportedEncodingException, MessagingException
	{

		InternetAddress fromAddressObj = new InternetAddress(fromAddress, "");
		InternetAddress toAddressObj = null;

		Address[] TO =null;
		//Message msg = new MimeMessage(session);
		/*
		 * if want to send multiple recipients 
		 */
		if (toAddress.contains(";") )
		{
			String[] toAddressary = toAddress.split(";");
			int toAddresslen = toAddressary.length;

			TO = new Address[toAddresslen];
			for ( int i=0;i<toAddresslen;i++)
			{
				TO[i] =  new InternetAddress(toAddressary[i]);
			}
			/* Address[] cc = new Address[] {
	                    ,
	                     new InternetAddress("sWER@gmail.com")};*/

			message.addRecipients(Message.RecipientType.TO,TO);
		}
		else
		{
			toAddressObj =new InternetAddress(toAddress, "");
			message.addRecipient(Message.RecipientType.TO,toAddressObj);
		}

		message.setFrom(fromAddressObj);
		//msg.setRecipients(Message.RecipientType.TO , "");

		message.setSubject(msgSubject);
		message.setSentDate(new Date());

		// creates message part
		MimeBodyPart messageBodyPart = new MimeBodyPart();
		messageBodyPart.setContent(msgBody, "text/html");

		// creates multi-part
		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(messageBodyPart);

		if (attachFiles != null && attachFiles[0]!=null && attachFiles.length > 0) 
		{
			for (String filePath : attachFiles) {
				MimeBodyPart attachPart = new MimeBodyPart();

				try {
					attachPart.attachFile(filePath);
				} catch (IOException ex) {
					ex.printStackTrace();
				}

				multipart.addBodyPart(attachPart);
			}
		}
		// sets the multi-part as e-mail's content
		message.setContent(multipart);

		t = session.getTransport("smtp");

		t.connect(EmailUser,EmailPassword);
		
		t.sendMessage(message, message.getAllRecipients());


	}


	/**
	 * Downloads new messages and saves attachments to disk if any.
	 * @param host
	 * @param port
	 * @param userName
	 * @param password
	 * @return 
	 * @throws IOException 
	 */
	public  Multimap<String, String> downloadEmailAttachments(String inboxFolder,String desFilePath) throws IOException {
		/*
		 * Create MultiValueMap 
		 */
		Multimap<String, String> multiMap = ArrayListMultimap.create();
		boolean attachementExist  = false;
		try {
			//Message msg = new MimeMessage(session);
			// connects to the message store

			//Message replyMessage = new MimeMessage(session);
			// opens the inbox folder
			Folder folderInbox = store.getFolder(inboxFolder);
			folderInbox.open(Folder.READ_WRITE);
			Flags seen = new Flags(Flags.Flag.SEEN);
			FlagTerm unseenFlagTerm = new FlagTerm(seen, false);

			// fetches new messages from server
			Message[] arrayMessages = folderInbox.search(unseenFlagTerm);

			for (int i = 0; i < arrayMessages.length; i++)
			{
				Message message = arrayMessages[i];

				this.fromAddress = message.getFrom();
				String from = fromAddress[0].toString();
				String subject = message.getSubject();
				String sentDate = message.getSentDate().toString();
				Address[] ccAddress = message.getRecipients(RecipientType.CC);
				String contentType = message.getContentType();
				String messageContent = "";
				String to =InternetAddress.toString(message.getRecipients(Message.RecipientType.TO));
				/*
				 * Check the mail subject pattern
				 * "Bethanaboiena, Venkata" <VBethanboien@mast.com> splilt<
				 */
				String  splitcc  =null;
				String   splitfrom  = from.split("<")[1].split(">")[0];
				splitcc  = splitfrom;
				if (ccAddress !=null)
				{
					for (int j=0 ; j< ccAddress.length ; j++)
					{
						splitcc   =splitcc+";"+ ccAddress[j].toString().split("<")[1].split(">")[0];
					}
				}
				if (checkSubjet(subject))
				{
					// store attachment file name, separated by comma
					String attachFiles = "";
					if (contentType.contains("multipart")) {
						// content may contain attachments
						Multipart multiPart = (Multipart) message.getContent();
						int numberOfParts = multiPart.getCount();
						for (int partCount = 0; partCount < numberOfParts; partCount++) {
							MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(partCount);
							if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
								// this part is attachment
								attachementExist =true;
								String fileName = part.getFileName();
								createDirectory("Reports\\"+GetDate());
								createDirectory("Reports\\"+GetDate()+"\\"+splitfrom);
								//Before saving the setting directory 

								setSaveDirectory(desFilePath);
								attachFiles += fileName + ", ";
								part.saveFile(saveDirectory + File.separator + fileName);
								multiMap.put(splitcc,desFilePath+"\\"+fileName);
							} else {
								// this part may be the message content
								messageContent = part.getContent().toString();
							}
						}
						if(!attachementExist)
						{
							replayMail( to, "Mail Does not have the attachement \n Thanks  ");
						}

					} else if (contentType.contains("text/plain") || contentType.contains("text/html")) {

						if (attachFiles.length() > 1) {
							attachFiles = attachFiles.substring(0, attachFiles.length() - 2);
						}
						else
						{
							if(!attachementExist)
							{
								replayMail( to, "Mail Does not have the attachement \n Thanks  ");
							}
						}
						Object content = message.getContent();
						if (content != null) {
							messageContent = content.toString();
						}
					}
				}
				else
				{
					System.out.println(message.isSet(Flag.SEEN));
					/*   t.connect(EmailUser, EmailPassword);
					   message.setFlag(Flag.SEEN, true);*/

				}
				/*print out details of each message
				System.out.println("Message #" + (i + 1) + ":");
				System.out.println("\t From: " + from);
				System.out.println("\t Subject: " + subject);
				System.out.println("\t Sent Date: " + sentDate);
				System.out.println("\t Message: " + messageContent);
				System.out.println("\t Attachments: " + attachFiles);*/
			}

			// disconnect
			folderInbox.close(false);
			store.close();
			return multiMap;
		} catch (NoSuchProviderException ex11) {
			System.out.println("No provider for pop3.");
			ex11.printStackTrace();

		} catch (MessagingException ex12) {
			System.out.println("Could not connect to the message store");
			ex12.printStackTrace();

		} catch (IOException ex13) {
			ex13.printStackTrace();

		}
		return multiMap;
	}
	public  void replayMail(String to,String msg) throws IOException, MessagingException
	{

		//	Message replyMessage = new MimeMessage(session);
		Message replyMessage = (MimeMessage)	 message.reply(true);
		replyMessage.setFrom(new InternetAddress(to));
		replyMessage.setRecipients(RecipientType.CC, message.getRecipients(RecipientType.CC));
		replyMessage.setText(msg);
		replyMessage.setReplyTo( message.getFrom());

		// t = session.getTransport("smtp");
		t.connect(EmailUser, EmailPassword);
		try {
			//connect to the smpt server using transport instance
			//change the user and password accordingly	

			t.sendMessage(replyMessage,
					fromAddress);
			replyMessage.setFlag(Flag.SEEN, true);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			// TODO: handle exception
		}
		finally {
			t.close();
		}
		System.out.println("message replied successfully ....");

	}
	public boolean checkSubjet(String Subject)
	{
		if(Subject.substring(0, SubjectPattern.length()).equals(SubjectPattern) || (Subject.equals("")) )
		{
			return true ;

		}
		return false;
	}
/*
//Runs this program with Gmail POP3 server
	public static void main(String[] args) throws IOException, MessagingException {
		String host = "webmail.limitedbrands.com";
			String port = "995";
			String userName = "TCOEAUTMaint@limited"; //username for the mail you want to read
			String password = "B3Side51"; //password
		 
		//String saveDirectory = "C:\\TestInbox\\Test";

		mailUtilites receiver = new mailUtilites();
		
		receiver.mailSessionInit("SAPConfig.properties");

		//receiver.setSaveDirectory(saveDirectory);
		receiver.mailSessionInit("Config.properties");
		//			mailUtilites.downloadEmailAttachments(host, port);
		receiver.downloadEmailAttachments();
	} */

	public String GetDate ()
	{
		int day, month, year;
		// int second, minute, hour;
		GregorianCalendar date = new GregorianCalendar();

		day = date.get(Calendar.DAY_OF_MONTH);
		month = date.get(Calendar.MONTH);
		year = date.get(Calendar.YEAR);

		/*    second = date.get(Calendar.SECOND);
			      minute = date.get(Calendar.MINUTE);t
			      hour = date.get(Calendar.HOUR);*/
		return Integer.toString(day)+Integer.toString(month+1)+Integer.toString(year);
	}
	public String GetDateTime ()
	{
		int day, month, year;
		int second, minute, hour;
		GregorianCalendar date = new GregorianCalendar();

		day = date.get(Calendar.DAY_OF_MONTH);
		month = date.get(Calendar.MONTH);
		year = date.get(Calendar.YEAR);

		second = date.get(Calendar.SECOND);
		minute = date.get(Calendar.MINUTE);
		hour = date.get(Calendar.HOUR);
		return Integer.toString(day)+Integer.toString(month+1)+Integer.toString(year)+Integer.toString(hour)+Integer.toString(minute)+Integer.toString(second);
	}
	public void createDirectory(String dirName)
	{
		File directory =null;
		if(dirName.contains(System.getProperty("user.dir")))
		{
			 directory = new File(dirName);
		}
		else
		{
		 directory = new File(System.getProperty("user.dir")+"\\"+dirName);
		}
		if (! directory.exists()){
			try {
				directory.mkdir();
			}catch(Exception ex)
			{
				System.out.print(ex.getMessage());
			}

		}
	}
	
	
	public static void getAllFiles(File dir, List<File> fileList) {
		try {
			File[] files = dir.listFiles();
			for (File file : files) {
				fileList.add(file);
				if (file.isDirectory()) {
					System.out.println("directory:" + file.getCanonicalPath());
					getAllFiles(file, fileList);
				} else {
					System.out.println("     file:" + file.getCanonicalPath());
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeZipFile(File directoryToZip, List<File> fileList) {

		try {
			FileOutputStream fos = new FileOutputStream(directoryToZip.getName() + ".zip");
			ZipOutputStream zos = new ZipOutputStream(fos);

			for (File file : fileList) {
				if (!file.isDirectory()) { // we only zip files, not directories
					addToZip(directoryToZip, file, zos);
				}
			}

			zos.close();
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void addToZip(File directoryToZip, File file, ZipOutputStream zos) throws FileNotFoundException,
			IOException {

		FileInputStream fis = new FileInputStream(file);

		// we want the zipEntry's path to be a relative path that is relative
		// to the directory being zipped, so chop off the rest of the path
		String zipFilePath = file.getCanonicalPath().substring(directoryToZip.getCanonicalPath().length() + 1,
				file.getCanonicalPath().length());
		System.out.println("Writing '" + zipFilePath + "' to zip file");
		ZipEntry zipEntry = new ZipEntry(zipFilePath);
		zos.putNextEntry(zipEntry);

		byte[] bytes = new byte[1024];
		int length;
		while ((length = fis.read(bytes)) >= 0) {
			zos.write(bytes, 0, length);
		}

		zos.closeEntry();
		fis.close();
	}
	//function adds/subtracts number of days from sysdate and return the date in required format.
		public  String addDaySystemDate(String format,int days) throws ParseException
		{
			Date date = new Date();
			Calendar c1 = Calendar.getInstance();
			c1.setTime(date); 

			java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat(format);
			c1.add(Calendar.DAY_OF_MONTH,days);
			return sdf.format(c1.getTime());
		}
		
		public boolean checkFileNotExists(String filepath)
		{
			
			Path fp = Paths.get(filepath);
			if (Files.notExists(fp))
				return true;
			else 
				return false;
		}
		
}
	

