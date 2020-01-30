Skype4B / Lync Certificate Checker Tool
=======================================

            

In many cases you may not have direct access to the other system you are connecting to in order to check whether the certificate it is using is valid, or has been signed by a trusted root certificate Authority. As a result, you may have issues connecting
 to the server and need to use complex tools like Wireshark to determine what the certificate being presented by the far end looks like. This can take time and involve installing software on servers, so I wanted to create a simple tool that doesn’t require
 any installation and can be run straight from a Powershell prompt. After doing some coding, that’s exactly what I created, introducing the Skype for Business Certificate Checker Tool:


![Image](https://github.com/jamescussen/skype4b-lync-certificate-checker-tool/raw/master/certchecker1.00new_sm.jpg) 


 


**Features:**


  *  Check the certificate being used by a server using the FQDN/IP and Port number of the server.

  *  Check the certificate of a Federation SRV record (_sipfederationtls._tcp.domain.com) simply by entering the SIP domain name and ticking the “FED SRV” checkbox.

  *  Check the SIP SRV record (_sip._tls.domain.com) by simply entering the SIP domain name and ticking the “SIP SRV” checkbox.


  *  Check the SIP SRV record (_sipinternaltls._tcp.domain.com) by simply entering the SIP domain name and ticking the “SIP INT SRV” checkbox.


  *  Select the DNS server you would like to use to resolve DNS from by entering a DNS Server IP address in the “DNS Server” field.

  *  “Show Advanced” checkbox will show all of the information in the certificate.

  *  The “Show Root Chain” will display the root certificate and all of the intermediate certificates that are applicable in the trust chain for the certificate.

  *  The “Test DNSLB Pool” checkbox is on by default and will instruct to the tool to test all of the IP Addresses that are resolved for a DNS Name. In the case of Skype for Business, we nearly always have multiple DNS records per A record for the
 purposes of DNS load balancing.  Rather than having to look all of the servers yourself, the tool will do this for you. Other servers in pool will be displayed in the Information text box in blue colour and will be tested
 directly via their IP Address rather than the DNS name. 
  *  Import multiple DNS name records from a CSV file. This is useful if you want to check a lot of servers in one sitting. See the “Import File Format” section for more details.

  *  Save certificate information out to a CSV file. This will save all of the certificate information out in table format that you can open in Excel for record keeping purposes.
**Note:** This export format is different than the one used in conjunction with the “Import” button.

  *  Comments section – The comments section will have information in it about things that may be wrong with the certificate to help you troubleshoot your issues.


 

Import File Format

You can import a CSV file containing many domains and servers to test if you choose (for example, this may be useful for checking a large list of federated domains). To do this you will first need to create a CSV file with all of the servers and/or domains
 that you want to test in it. The format of the CSV for each of the record types will look like:


**Header row:** Domain,Type,Port


**Example Federation Record:** 'microsoft.com','FED','',


**Example SIP Record:** 'microsoft.com','SIP','',


**Example Internal SIP Record:** 'microsoft.com','SIPINT','',


**Example direct Record:** 'sip.microsoft.com','DIR','5061',


** **


**For full details on the Tool please refer to the blog post here: **


**[http://www.myteamslab.com/2016/12/skype4b-lync-certificate-checker-tool.html](http://www.myteamslab.com/2016/12/skype4b-lync-certificate-checker-tool.html)**


 






        
    
