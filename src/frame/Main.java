package frame;


import java.awt.BorderLayout;

import javax.swing.JDialog;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JFrame;
import javax.swing.JLabel;
import java.awt.Rectangle;
import javax.swing.JTextField;
import javax.swing.JButton;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.scribe.builder.ServiceBuilder;
import org.scribe.builder.api.LinkedInApi;
import org.scribe.model.OAuthRequest;
import org.scribe.model.Response;
import org.scribe.model.Token;
import org.scribe.model.Verb;
import org.scribe.model.Verifier;
import org.scribe.oauth.OAuthService;

import unti.AuthHandler;
import unti.XLSUnti;

import java.awt.Dimension;
import java.awt.Point;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.Iterator;
import javax.swing.JPasswordField;
import javax.swing.JProgressBar;

public class Main extends JFrame {

	private static final long serialVersionUID = 1L;
	private JPanel jContentPane = null;
	private JButton jButton = null;
	private JLabel jLabel1 = null;
	private JButton jButton1 = null;
	private JTextField jTextField1 = null;
	private JLabel jLabel2 = null;
	private String url="";  //  @jve:decl-index=0:
	private  OAuthService service = new ServiceBuilder()
    .provider(LinkedInApi.class)
    .apiKey("a2f5jedymg68")
    .apiSecret("heJkIFSpkO2ZVjlo")
    .build();//  @jve:decl-index=0:
	private Token requestToken = service.getRequestToken();  //  @jve:decl-index=0:
	private Token accessToken = null;  //  @jve:decl-index=0:
	private String op="";
	private boolean flag;
	private JButton jButton2 = null;
	File file = new File("service.dat");  //  @jve:decl-index=0:
	private JProgressBar jProgressBar = null;
	/**
	 * This is the default constructor
	 * @throws ClassNotFoundException 
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 */
	public Main() throws  IOException, ClassNotFoundException {
		super();
		initialize();
	}

	/**
	 * This method initializes this
	 * 
	 * @return void
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 * @throws ClassNotFoundException 
	 */
	private void initialize() throws IOException, ClassNotFoundException {
		this.setSize(300, 258);
		this.setContentPane(getJContentPane());
		this.setTitle("LinkedIn Friends Output");
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		Center(this);
        if(file.exists()){
            //if the file exists we assume it has the AuthHandler in it - which in turn contains the Access Token
            ObjectInputStream inputStream = new ObjectInputStream(new FileInputStream(file));
            AuthHandler authHandler = (AuthHandler) inputStream.readObject();
            accessToken = authHandler.getAccessToken();
            jLabel1.show(false);
        	jLabel2.show(false);
        	getJButton1().show(false);
        	getJTextField1().show(false);
        	getJButton().setLocation(150, 5);
        	setSize(265, 120);
        	getJProgressBar().setSize(240, 22);
        	getJProgressBar().setLocation(5, 45);
        	
            flag=true;
        } else {
        	flag=false;
        	getJButton2().show(false);
        }
        
		
		
		
		
		
		
	}

	/**
	 * This method initializes jContentPane
	 * 
	 * @return javax.swing.JPanel
	 */
	private JPanel getJContentPane() {
		if (jContentPane == null) {
			jLabel2 = new JLabel();
			jLabel2.setBounds(new Rectangle(13, 98, 177, 27));
			jLabel2.setText("Second Step:enter the code:");
			jLabel1 = new JLabel();
			jLabel1.setBounds(new Rectangle(12, 11, 261, 33));
			jLabel1.setText("First Step: Access the security code");
			jContentPane = new JPanel();
			jContentPane.setLayout(null);
			jContentPane.add(getJButton(), null);
			jContentPane.add(jLabel1, null);
			jContentPane.add(getJButton1(), null);
			jContentPane.add(getJTextField1(), null);
			jContentPane.add(jLabel2, null);
			jContentPane.add(getJButton2(), null);
			jContentPane.add(getJProgressBar(), null);
		}
		return jContentPane;
	}

	/**
	 * This method initializes jButton	
	 * 	
	 * @return javax.swing.JButton	
	 */
	private JButton getJButton() {
		if (jButton == null) {
			jButton = new JButton("Start");
			jButton.setSize(85, 30);
			jButton.setLocation(103, 134);	
			jButton.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					if(!flag){
						Verifier verifier = new Verifier(getJTextField1().getText());
						try {
							accessToken = service.getAccessToken(requestToken,verifier);
							AuthHandler authHandler = new AuthHandler(accessToken);
							
							ObjectOutputStream outputStream = new ObjectOutputStream(new FileOutputStream("service.dat"));
							outputStream.writeObject(authHandler);
							outputStream.close();
							accessToken = authHandler.getAccessToken();
						} catch (Exception e2) {
							JOptionPane.showMessageDialog(getJContentPane(),"security code error!","error",JOptionPane.ERROR_MESSAGE);
						}	
					}
					
					
	               String url = "http://api.linkedin.com/v1/people/~/connections:(email-address,distance,member-url-resources,location,id,first-name,last-name,headline,industry,num-connections,honors,positions,publications,languages,skills,certifications,educations,phone-numbers,date-of-birth,main-address,picture-url,public-profile-url)";
	    	        //String url = "http://api.linkedin.com/v1/people/id=EsJguu8BtB:(id,first-name,last-name,headline,industry,num-connections,honors,positions,publications,languages,skills,certifications,educations,phone-numbers,date-of-birth,main-address,picture-url,public-profile-url)";

	    	       OAuthRequest request = new OAuthRequest(Verb.GET, url);
	    	       request.addHeader("x-li-format", "json");
	    	       service.signRequest(accessToken, request);
	    	       Response response = request.send();
	    	       System.out.println(response.getBody());
	    	       JSONObject j1 = JSONObject.fromObject(response.getBody());
	    	       JSONArray j2 = JSONArray.fromObject(j1.get("values"));
	    	       getJProgressBar().setOrientation(JProgressBar.HORIZONTAL);
	    	       getJProgressBar().setMinimum(0);
	    	       getJProgressBar().setMaximum(Integer.valueOf(j1.get("_total").toString()));
	    	      
	    	       String path = "1.xls";
	    	       FileOutputStream fos;
				try {
					fos = new FileOutputStream(path);
				
	    	       HSSFWorkbook workbook =  new  HSSFWorkbook();
	    	       HSSFCellStyle style = workbook.createCellStyle();
	    	       HSSFCellStyle leftstyle = workbook.createCellStyle();
	    	       HSSFCellStyle rightstyle = workbook.createCellStyle();
	    	       style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	    	       leftstyle.setBorderLeft((short)2); 
	    	       rightstyle.setBorderRight((short)2);
	    	       HSSFSheet sheet = workbook.createSheet();
	    	       
	    	       for(int i=0;i<66;i++){
	    	       		if(i==6)continue;
	    	       		if(i==7)continue;
	    	       		sheet.setDefaultColumnStyle(i, style);
	    	       }
	    	       XLSUnti.createXLSTitle(sheet,leftstyle,rightstyle);
	    	       
	    	       getJProgressBar().setValue(0);
	    	       
	    	       Iterator<JSONObject> ij = j2.iterator();
	    	       int iid=1;
	    	       while(ij.hasNext()){
	    	    	   getJProgressBar().setValue(iid);
	    	    	   JSONObject result = ij.next();
	    	    	   String[] s = new String[66];
	    	    	   s[0] = String.valueOf(iid);
	    	    	   if(result.get("firstName")!=null)s[1]=(String) result.get("firstName");
	    	    	   if(result.get("lastName")!=null)s[2]=(String) result.get("lastName");
	    	    	   if(result.get("headline")!=null)s[3]=(String) result.get("headline");
	    	    	   if(result.get("industry")!=null)s[4]=(String) result.get("industry");
	    	    	   if(result.get("numConnections")!=null)s[5]=String.valueOf(result.get("numConnections"));
	    	    	   if(result.get("pictureUrl")!=null)s[6]=(String) result.get("pictureUrl");
	    	    	   if(result.get("publicProfileUrl")!=null)s[7]=(String) result.get("publicProfileUrl");

	    	    	   JSONArray positionresult = null;
	    	    	   if(result.get("positions")!=null&&JSONObject.fromObject(result.get("positions")).get("values")!=null){
	    	        		positionresult = JSONObject.fromObject(result.get("positions")).getJSONArray("values");
	    	        		Iterator<JSONObject> ip = positionresult.iterator();
	    	        		int i=0;
	    	        		while(ip.hasNext()){
	    	        			JSONObject pr = ip.next();
	    	        			if(pr.get("company")!=null){
	    	        				JSONObject companyresult = null;
	    	        				companyresult = (JSONObject) pr.get("company");
	    	        				if(companyresult.get("id")!=null)s[16+i*13]=String.valueOf(companyresult.get("id"));
	    	        				if(companyresult.get("industry")!=null)s[19+i*13]=String.valueOf(companyresult.get("industry"));
	    	        				if(companyresult.get("name")!=null)s[17+i*13]=String.valueOf(companyresult.get("name"));
	    	        				if(companyresult.get("size")!=null)s[20+i*13]=String.valueOf(companyresult.get("size"));
	    	        				if(companyresult.get("type")!=null)s[18+i*13]=String.valueOf(companyresult.get("type"));
	    	        			}
	    	        			if(pr.get("startDate")!=null){
	    	        				JSONObject sresult = null;
	    	        				sresult = (JSONObject) pr.get("startDate");
	    	        				if(sresult.get("year")!=null)s[11+i*13]=String.valueOf(sresult.get("year"));
	    	        				if(sresult.get("month")!=null)s[12+i*13]=String.valueOf(sresult.get("month"));
	    	        			}
	    	        			if(pr.get("endDate")!=null){
	    	        				JSONObject eresult = null;
	    	        				eresult = (JSONObject) pr.get("endDate");
	    	        				if(eresult.get("year")!=null)s[14+i*13]=String.valueOf(eresult.get("year"));
	    	        				if(eresult.get("month")!=null)s[13+i*13]=String.valueOf(eresult.get("month"));
	    	        			}
	    	        			if(pr.get("isCurrent")!=null)s[15+i*13]=String.valueOf(pr.get("isCurrent"));
	    	        			if(pr.get("summary")!=null)s[10+i*13]=(String) pr.get("summary");
	    	        			if(pr.get("title")!=null)s[9+i*13]=(String) pr.get("title");
	    	        			if(pr.get("id")!=null)s[8+i*13]=String.valueOf(pr.get("id"));
	    	        			
	    	        			if(i==3){
	    	        				break;
	    	        			}else{
	    	        				i++;
	    	        			}
	    	        			
	    	        		}
	    	        	}
	    	    	   
	    	    	   
	    	    	   
	    	    	   JSONArray educationresult = null;
	    	    	   if(result.get("educations")!=null&&JSONObject.fromObject(result.get("educations")).get("values")!=null){
	    	        		educationresult = JSONObject.fromObject(result.get("educations")).getJSONArray("values");
	    	        		Iterator<JSONObject> ie = educationresult.iterator();
	    	        		
	    	        		int i=0;
	    	        		while(ie.hasNext()){
	    	        			JSONObject er = ie.next();
	    	        			if(er.get("degree")!=null)s[47+i*5]=String.valueOf(er.get("degree"));
	    	        			if(er.get("activities")!=null)
	    	        			if(er.get("fieldOfStudy")!=null)s[50+i*5]=String.valueOf(er.get("fieldOfStudy"));
	    	        			if(er.get("notes")!=null)
	    	        			if(er.get("schoolName")!=null)s[51+i*5]=String.valueOf(er.get("schoolName"));
	    	        			if(er.get("startDate")!=null){
	    	        				JSONObject sresult = null;
	    	        				sresult = (JSONObject) er.get("startDate");
	    	        				if(sresult.get("year")!=null)s[48+i*5]=String.valueOf(sresult.get("year"));
	    	        			}
	    	        			if(er.get("endDate")!=null){
	    	        				JSONObject eresult = null;
	    	        				eresult = (JSONObject) er.get("endDate");
	    	        				if(eresult.get("year")!=null)s[49+i*5]=String.valueOf(eresult.get("year"));
	    	        			}
	    	        			if(i==3){
	    	        				break;
	    	        			}else{
	    	        				i++;
	    	        			}
	    	        		}
	    	        		
	    	        	}
	    	    	   
	    	        		
	    	    	   	
	    	        		
	    	        		
	    	    	   JSONArray phoneresult = null;
	    	        	if(result.get("phoneNumbers")!=null&&JSONObject.fromObject(result.get("phoneNumbers")).get("values")!=null){
	    	        		phoneresult = JSONObject.fromObject(result.get("phoneNumbers")).getJSONArray("values");
	    	        		Iterator<JSONObject> ip = phoneresult.iterator();
	    	        		
	    	        		while(ip.hasNext()){
	    	        			JSONObject pr = ip.next();
	    	        			if(pr.get("phoneType")!=null)
	    	        			if(pr.get("phoneNumber")!=null)s[62]=String.valueOf(pr.get("phoneNumber"));
	    	        		}
	    	        	}
	    	    	   
	    	    	   
	    	        	JSONArray skillresult = null;
	    	        	if(result.get("skills")!=null&&JSONObject.fromObject(result.get("skills")).get("values")!=null){
	    	        		skillresult = JSONObject.fromObject(result.get("skills")).getJSONArray("values");
	    	        		Iterator<JSONObject> is = skillresult.iterator();
	    	        		s[63]="";
	    	        		while(is.hasNext()){
	    	        			JSONObject sr = is.next();
	    	        			if(sr.get("skill")!=null){
	    	        				s[63] += JSONObject.fromObject(sr.get("skill")).get("name") + " | ";
	    	        			}
	    	        		}
	    	        	}
	    	    	   
	    	        	
	    	        	JSONObject locationresult = null;
	    	        	if(result.get("location")!=null&&JSONObject.fromObject(result.get("location")).get("name")!=null){
	    	        		locationresult = JSONObject.fromObject(result.get("location"));
	    	        		s[64] = (String) locationresult.get("name");
	    	        	}
	    	        	
	    	        	
	    	        	
	    	        	JSONArray memberUrlResourcesresult = null;
	    	        	if(result.get("memberUrlResources")!=null&&JSONObject.fromObject(result.get("memberUrlResources")).get("values")!=null){
	    	        		memberUrlResourcesresult = JSONObject.fromObject(result.get("memberUrlResources")).getJSONArray("values");
	    	        		Iterator<JSONObject> is = memberUrlResourcesresult.iterator();
	    	        		while(is.hasNext()){
	    	        			JSONObject sr = is.next();
	    	        			
	    	        			if(sr.get("url")!=null){
	    	        				s[65] = (String) sr.get("url");
	    	        			}
	    	        		}
	    	        	}
	    	        	
	    	    	   
	    	        	sheet.setColumnWidth(0,1000);
	    	        	sheet.setColumnWidth(3,8000);
	    	        	sheet.setColumnWidth(4,5000);
	    	        	sheet.setColumnWidth(5,4000);
	    	        	sheet.setColumnWidth(6,7000);
	    	        	sheet.setColumnWidth(7,7000);
	    	        	
	    	        	for (int i = 0; i < 3; i++) {
							sheet.setColumnWidth(8+i*13, 4000);
							sheet.setColumnWidth(9+i*13, 8500);
							sheet.setColumnWidth(10+i*13, 6500);
							sheet.setColumnWidth(11+i*13, 2000);
							sheet.setColumnWidth(12+i*13, 2500);
							sheet.setColumnWidth(13+i*13, 2000);
							sheet.setColumnWidth(14+i*13, 2000);
							sheet.setColumnWidth(15+i*13, 2000);
							sheet.setColumnWidth(16+i*13, 3000);
							sheet.setColumnWidth(17+i*13, 7500);
							sheet.setColumnWidth(18+i*13, 5000);
							sheet.setColumnWidth(19+i*13, 5500);
							sheet.setColumnWidth(20+i*13, 5000);
						}
	    	        	
						for (int i = 0; i < 3; i++) {
							sheet.setColumnWidth(47+i*5, 6500);
							sheet.setColumnWidth(48+i*5, 2000);
							sheet.setColumnWidth(49+i*5, 2000);
							sheet.setColumnWidth(50+i*5, 5000);
							sheet.setColumnWidth(51+i*5, 10000);
						}
						sheet.setColumnWidth(62,4000);
	    	        	sheet.setColumnWidth(63,5000);
	    	        	sheet.setColumnWidth(64,2000);
	    	        	sheet.setColumnWidth(65,8000);
	    	        	
	    	        	HSSFRow row = sheet.createRow(++iid);
	    	        	XLSUnti.insertOneData(row,s,leftstyle,rightstyle);
	    	    	  
	    	       }
	    	       	   

						workbook.write(fos);
						fos.close();
						JOptionPane.showMessageDialog(getJContentPane(),"finish!","success",JOptionPane.INFORMATION_MESSAGE);
						getJProgressBar().setValue(0);
				} catch (FileNotFoundException e2) {
						// TODO Auto-generated catch block
						e2.printStackTrace();
						JOptionPane.showMessageDialog(getJContentPane(),"Another program is using the xls!","error",JOptionPane.ERROR_MESSAGE);
						
				} catch (IOException e1) {
						// TODO Auto-generated catch block
					e1.printStackTrace();
					JOptionPane.showMessageDialog(getJContentPane(),"Unknown error!","error",JOptionPane.ERROR_MESSAGE);
						
				} catch(Exception e12){
					e12.printStackTrace();
					JOptionPane.showMessageDialog(getJContentPane(),"Unknown error!","error",JOptionPane.ERROR_MESSAGE);
						
				}
	    	       
				}
			});
		}
		return jButton;
	}

	/**
	 * This method initializes jButton1	
	 * 	
	 * @return javax.swing.JButton	
	 */
	private JButton getJButton1() {
		if (jButton1 == null) {
			jButton1 = new JButton("Access code");
			jButton1.setBounds(new Rectangle(82, 55, 120, 28));
			jButton1.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					
					System.out.println(service.getAuthorizationUrl(requestToken));
					try {
						Runtime.getRuntime().exec("rundll32 url.dll,FileProtocolHandler " + service.getAuthorizationUrl(requestToken));
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			});
		}
		return jButton1;
	}

	/**
	 * This method initializes jTextField1	
	 * 	
	 * @return javax.swing.JTextField	
	 */
	private JTextField getJTextField1() {
		if (jTextField1 == null) {
			jTextField1 = new JTextField();
			jTextField1.setBounds(new Rectangle(179, 98, 83, 27));
		}
		return jTextField1;
	}
	
	/**
	 * This method initializes jButton2	
	 * 	
	 * @return javax.swing.JButton	
	 */
	private JButton getJButton2() {
		if (jButton2 == null) {
			jButton2 = new JButton("Logout");
			jButton2.setBounds(new Rectangle(14, 5, 82, 28));
			jButton2.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					jLabel1.show(true);
		        	jLabel2.show(true);
		        	getJButton1().show(true);
		        	getJButton2().show(false);
		        	getJTextField1().show(true);
		        	getJButton().setLocation(103, 134);
		        	getJProgressBar().setBounds(new Rectangle(10, 184, 266, 22));
		        	setSize(300, 258);
		        	file.delete();
				}
			});
		}
		return jButton2;
	}

	
	public static void Center(JFrame frame){ 
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize(); 
		Dimension frameSize = frame.getSize(); 
		if (frameSize.height > screenSize.height) { 
			frameSize.height = screenSize.height; 
		} 
		if (frameSize.width > screenSize.width) { 
			frameSize.width = screenSize.width; 
		}
		
		frame.setLocation((screenSize.width - frameSize.width) / 2, (screenSize.height - frameSize.height) / 2); 
		frame.show(); 
		
	} 
	
	
	
	/**
	 * This method initializes jProgressBar	
	 * 	
	 * @return javax.swing.JProgressBar	
	 */
	private JProgressBar getJProgressBar() {
		if (jProgressBar == null) {
			jProgressBar = new JProgressBar();
			jProgressBar.setBounds(new Rectangle(10, 184, 266, 22));
		}
		return jProgressBar;
	}

	public static void main(String[] args) throws IOException, ClassNotFoundException {
		new Main().show();
		// TODO Auto-generated method stub
	}

}  //  @jve:decl-index=0:visual-constraint="392,24"
