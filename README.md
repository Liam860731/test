package tw.com.prm;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeUtility;
import javax.mail.internet.ParseException;

import tw.com.synergy.cynosure.util.ServiceUtil;
import tw.com.synergy.cynosure.util.TransFormat;
import tw.com.synergy.mesware.util.CheckUtil;
import tw.com.synergy.mesware.util.CynosureServiceDescription;
import tw.com.synergy.mesware.util.LogicalException;
import tw.com.synergy.mesware.util.LogicalService;
import tw.com.synergy.mesware.util.PersistenceException;
import tw.com.synergy.mesware.util.SqlBean;

import com.sun.mail.imap.IMAPStore;

/**
 * 
 * <p>Copyright 2006 All Rights Reserved by Synergy.</p>
 * <p>Modification : (Date-Version-Author-Description)</p>
 * <p>----------------------------------------------------------------------</p>
 * <p>2022.10.25 - 1.0.0 - x06144 - Setting Up class.</p>
 * @servicedoc
 */

@CynosureServiceDescription("")
public class ReseiveSpcScheduleMail extends LogicalService
{
	private InputSpo _inSpo;
    private OutputSpo _outSpo = new OutputSpo();

	/**
	 * Business logic.
	 * @throws LogicalException le
     * @throws PersistenceException pe
     * @throws Exception e
	 */
    public void excute() throws LogicalException, PersistenceException, Exception
    {
    	reseive();
    	// 更新AutoMonitorLog時間
    	int updateCnt = updSysAttribData(); 
    }
    
    private void reseive( ) throws LogicalException, PersistenceException, Exception
	{    	
    	    	
    	String host = ServiceUtil.GetLogicalDescrip(this._conn, _log, "SysMail","SMTPServer");
    	String spc_monitor_user = ServiceUtil.GetLogicalDescrip(this._conn, _log, "SysMail","spc_monitor_user");
    	String spc_monitor_password = ServiceUtil.GetLogicalDescrip(this._conn, _log, "SysMail","spc_monitor_password");
    	
        // Create a Properties object to contain connection configuration information.
//    	Properties props = new Properties();
//        props.setProperty("mail.transport.protocol", "smtp");
//        props.setProperty("mail.smtp.port", "25");
//
//        // Set properties indicating that we want to use STARTTLS to encrypt the connection.
//        // The SMTP session will begin on an unencrypted connection, and then the client
//        // will issue a STARTTLS command to upgrade to an encrypted connection.
//        props.setProperty("mail.smtp.auth", "true");
//        props.setProperty("mail.smtp.starttls.enable", "true");
//        props.setProperty("mail.smtp.starttls.required", "true");
//        props.setProperty("mail.smtp.ssl.enable", "false");
//        props.setProperty("mail.debug", "true");
    	
    	//Properties properties
    	Properties props = new Properties();
    	props.setProperty("mail.imap.host",host);
    	props.setProperty("mail.imap.port","143");
    	props.setProperty("mail.store.protocol","imap");    	    
    	props.setProperty("mail.debug.auth", "true");    
    	props.setProperty("mail.imap.starttls.enable", "true");
    	props.setProperty("mail.debug", "true");
    	    	 
    	Session session = Session.getInstance(props);    	
    	IMAPStore store = (IMAPStore) session.getStore("imap");
    	store.connect(host,spc_monitor_user, spc_monitor_password);
    	
        Folder folder = store.getFolder("INBOX");  
        folder.open(Folder.READ_WRITE);
        Message[] messages = folder.getMessages();  
        if(messages.length > 0)
        {
        	parseMessage(messages);  
        }	
        folder.close(true);  
        store.close();     	    	    	
	}
    
    public  void parseMessage(Message ...messages) throws LogicalException,MessagingException, IOException, ParseException, PersistenceException, Exception
    {  
         
        String sScheduleName = "";
        String sGroupName = "";
        String sExcelFile = "";
        String sSpecMax = "";
		String sSpecMin = "";
        String sExcelFileName = "";
        Double sExcelFileSize = 0.0;
        String sSentDate = "";
        
        _log.debug("messages.length  = " + messages.length ) ;
        String sSubject = ""; 

        for (int i = 0, count = messages.length; i < count; i++) 
        {  
        	_log.debug("count  = " + i ) ;
            MimeMessage msg = (MimeMessage) messages[i];  
           _log.debug("------------------Parse Message" + msg.getMessageNumber() + "Mail Start-------------------- ");  
             
           sSubject = getSubject(msg);
           _log.debug("Subject: " + sSubject);
           
           _log.debug("From: " + getFrom(msg));  
           sSentDate = getSentDate(msg, "yyyy/MM/dd HH:mm:ss");
           _log.debug("Sent Date：" + sSentDate);

            StringBuffer content = new StringBuffer(30);  
            getMailTextContent(msg, content);  
           _log.debug("Mail Text Content：" + (content));  
           _log.debug("------------------Parse Message =" + msg.getMessageNumber() + "= Mail End-------------------- ");  
           
           if(sSubject.indexOf("[SPC]") > -1 && sSubject.indexOf("未傳遞的主旨") < 0 )
           {
        	   _log.debug("[SPC] Start------------------------------------------------------------");
        	   _log.debug("sSubject = " + sSubject );        	   
        	   sScheduleName = sSubject.split("\\[SPC\\]")[1].toString().trim();
//        	   List<Map> monitorChartMappingList = ServiceUtil.getMESSysParam(_conn,_log, "MonitorChartMapping", sScheduleName);
        	   List<Map> monitorChartMappingList =  ServiceUtil.getMESSysAttribData(_conn, _log,"monitor_performance_job","SPC",sScheduleName+"%");
        	   if (!CheckUtil.isNull(monitorChartMappingList)) 
			   {
//        		   sScheduleName = TransFormat.getProperty(monitorChartMappingList.get(0),"logicaldescrip", "");
//        		   sGroupName = TransFormat.getProperty(monitorChartMappingList.get(0),"memo", "");
					sScheduleName = TransFormat.getProperty(monitorChartMappingList.get(0), "attribval2", "");
					sGroupName = TransFormat.getProperty(monitorChartMappingList.get(0), "attribval6", "");
					sSpecMax = TransFormat.getProperty(monitorChartMappingList.get(0), "attribval8", "");
					sSpecMin = TransFormat.getProperty(monitorChartMappingList.get(0), "attribval9", "");
			   }
        	   _log.debug("sScheduleName = " + sScheduleName );
        	   String sEndDateTime = "";
        	   String tmpEndDateTime = content.toString().split("End DateTime")[1].toString().split("<br>")[0].toString().trim();
        	   _log.debug("tmpEndDateTime = " + tmpEndDateTime);
        	   sEndDateTime = tmpEndDateTime.replaceFirst(":", "");
        	   _log.debug("EndDateTime = " + sEndDateTime );
        	   
	        	String tmpDay = "";
	        	String sReceiveDate = "";
	           	Date dReceiveDate = msg.getSentDate();
	       		Calendar calendar = Calendar.getInstance();
	       		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	       		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy/MM/dd");
	       		calendar.setTime(dReceiveDate);
	       		sReceiveDate = sdf.format(dReceiveDate);
	       		_log.debug("sReceiveDate = " + sReceiveDate);
	       		//2022/11/01 06:00:00
	       		Date dEndDateTime = sdf.parse(sEndDateTime);
	       		long execTime = (dReceiveDate.getTime() - dEndDateTime.getTime())/(1000*60);
	       		
	       		if(execTime < 1)
	       		{
	       			execTime = 1;
	       		}	
	           	_log.debug("execTime = " +  execTime);
       		
	           	sExcelFile = content.toString().split("<a href=\"")[1].toString().split("\">")[0].toString().trim();
	           	sExcelFile = sExcelFile.replace("file:///", "").replace("%20", " ");
        	   
        	   _log.debug("sExcelFile = " + sExcelFile );
        	   File myObj = new File(sExcelFile);
        	   if (myObj.exists()) 
        	   {
        		   sExcelFileName = myObj.getName();
        		   sExcelFileSize =  (double) (myObj.length() / 1024) ;
        		   
					_log.debug("File name: " + sExcelFileName);
					_log.debug("File size in KB " + sExcelFileSize);
        	    	
        	   }
        	   insertSpcMonitorLog(sScheduleName,sEndDateTime,sReceiveDate,execTime,sExcelFileName,sExcelFileSize,"SPC",sGroupName,sSpecMax,sSpecMin);
        	   Message message = messages[i];
               String subject = message.getSubject();
//                set the DELETE flag to true
//               message.setFlag(Flags.Flag.DELETED, true);
              _log.debug("Marked DELETE for message : " + sSubject);  
        	  _log.debug("[SPC]  END------------------------------------------------------------");
           }             
        }
    } 
    
    public static String getSubject(MimeMessage msg) throws UnsupportedEncodingException, MessagingException 
    {  
        return MimeUtility.decodeText(msg.getSubject());  
    } 
    

    public static String getFrom(MimeMessage msg) throws MessagingException, UnsupportedEncodingException 
    {  
        String from = "";  
        Address[] froms = msg.getFrom();  

        InternetAddress address = (InternetAddress) froms[0];  
        String person = address.getPersonal();  
        if (person != null) {  
            person = MimeUtility.decodeText(person) + " ";  
        } else {  
            person = "";  
        }  
        from = person + "<" + address.getAddress() + ">";  
          
        return from;  
    }
    
    	
    public static void getMailTextContent(Part part, StringBuffer content) throws MessagingException, IOException 
    {  
        boolean isContainTextAttach = part.getContentType().indexOf("name") > 0;   
        if (part.isMimeType("text/*") && !isContainTextAttach) {  
            content.append(part.getContent().toString().trim());  
        } else if (part.isMimeType("message/rfc822")) {   
            getMailTextContent((Part)part.getContent(),content);  
        } else if (part.isMimeType("multipart/*")) {  
            Multipart multipart = (Multipart) part.getContent();  
            int partCount = multipart.getCount();  
            for (int i = 0; i < partCount; i++) {  
                BodyPart bodyPart = multipart.getBodyPart(i);  
                getMailTextContent(bodyPart,content);  
            }  
        }  
    } 
               
    public static String getSentDate(MimeMessage msg, String pattern) throws MessagingException {  
        Date receivedDate = msg.getSentDate();  
        if (receivedDate == null)  
            return "";  
          
        
        return new SimpleDateFormat(pattern).format(receivedDate);  
    }
    
    /**
	 * Jackie Add For Monitor Job is working
	 * @param conn
	 * @return
	 * @throws LogicalException
	 * @throws PersistenceException
	 * @throws Exception
	 */
	private int updSysAttribData() throws LogicalException, PersistenceException, Exception 
	{		

		StringBuffer sql = new StringBuffer();
		sql = new StringBuffer();
		sql.append(" UPDATE tblsysattribdata SET attribval5 = to_char(sysdate,'yyyy/mm/dd hh24:mi:ss') WHERE attribtype LIKE 'AutoMonitorLog%' AND attribval1 = 'ReseiveSpcScheduleMail' ");
		
		SqlBean sqlBean = new SqlBean(_conn, sql.toString(), _log);
		return sqlBean.executeUpdate();
	}
	
	 private void insertSpcMonitorLog(String sScheduleName, String sActiveDate,
	    		String sReceiveDate, long execTime, String sExcelFileName,Double sExcelFileSize, String sSystem,String sChartGroup,String max,String min)  throws LogicalException, PersistenceException, Exception
	    {
	    	
	    	StringBuffer sql = new StringBuffer();
			sql.append("    INSERT INTO H_PRM_MNTR (name,start_time,end_time,exec_time,file_name,file_size,lm_user,lm_time, system,chart_group,max,min) ");
			sql.append("    VALUES (?, TO_DATE(?,'yyyy/MM/dd HH24:mi:ss'), TO_DATE(?,'yyyy/MM/dd HH24:mi:ss') , ?, ?, ?, ?,sysdate, ?, ?, ?, ?) ");

			SqlBean sqlBean = new SqlBean(_conn,sql.toString(),_log);
			sqlBean.addParameter(sScheduleName);
			sqlBean.addParameter(sActiveDate);
			sqlBean.addParameter(sReceiveDate);
			sqlBean.addParameter(execTime);
			sqlBean.addParameter(sExcelFileName);
			sqlBean.addParameter(sExcelFileSize);
			sqlBean.addParameter(this.getUserID());
			sqlBean.addParameter(sSystem);
			sqlBean.addParameter(sChartGroup);
			sqlBean.addParameter(max);
			sqlBean.addParameter(min);
			sqlBean.executeUpdate();		
		}
       
	/**
     * validateParameter.
     * @throws LogicalException le
     */
    public void validateParameter() throws LogicalException
    {

    }
    
	/**
     * set input spo
     */
    public void setInput(InputSpo in)
    {
        _inSpo = in;        
    }
    
    /**
     * get output spo
     */
    public OutputSpo getOutput()
    {
        return _outSpo;
    }
    
	/**
	 * Input Service Parameter Object 
	 */
	public static class InputSpo
    {

	}
	
	/**
     * Output Service Parameter Object
     */
    public static class OutputSpo
    {

    }	
}

