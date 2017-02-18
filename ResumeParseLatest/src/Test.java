import java.util.*;
import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;
public class Test {

	 public static void main(String[] args) {
	        // Recipient's email ID needs to be mentioned.
	          String to = "jayadeep66@gmail.com";

	          // Sender's email ID needs to be mentioned
	          String from = "jayadeep66@gmail.com";


	          // Get system properties
	          Properties properties = System.getProperties();

	          properties.put("mail.smtp.starttls.enable", "true"); 
	          properties.put("mail.smtp.host", "smtp.gmail.com");

	          properties.put("mail.smtp.port", "587");
	          properties.put("mail.smtp.auth", "true");
	          Authenticator authenticator = new Authenticator () {
	                public PasswordAuthentication getPasswordAuthentication(){
	                    return new PasswordAuthentication("jayadeep66@gmail.com","jagatdeep");//userid and password for "from" email address 
	                }
	            };

	            Session session = Session.getDefaultInstance( properties , authenticator);  
	          try{
	             // Create a default MimeMessage object.
	             MimeMessage message = new MimeMessage(session);

	             // Set From: header field of the header.
	             message.setFrom(new InternetAddress(from));

	             // Set To: header field of the header.
	             message.addRecipient(Message.RecipientType.TO,
	                                      new InternetAddress(to));

	             // Set Subject: header field
	             message.setSubject("This is the Subject Line!");

	             // Now set the actual message
	             message.setText("This is actual message");

	             // Send message
	             Transport.send(message);
	             System.out.println("Sent message successfully....");
	          }catch (MessagingException mex) {
	             mex.printStackTrace();
	          }
	    }

	}