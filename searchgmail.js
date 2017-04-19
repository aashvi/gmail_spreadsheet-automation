function sales2() {

  var pay_initSpreadSheet = SpreadsheetApp.openById("1IZvScf_KAiw4hYf0U0cKzvE-NfFlYxyZKrAgvj4W4KE");
    var pay_sucSpreadSheet = SpreadsheetApp.openById("1KOku9IvtSYCKdXbFgJIJ3c2y8r0GoerUNRTvKMvVnjw");
  var pay_fSpreadSheet = SpreadsheetApp.openById("11M6NtXi_pcDxFQbXIczU5_jrSVfme6gbNqyUpMA7Kj0");
  var tckt_issSpreadSheet = SpreadsheetApp.openById("1FtVIyxrAmlfn5wxfsxSqQaL41IaWTf5pPBQ8PePl4ww");
    var offlineSpreadSheet = SpreadsheetApp.openById("1fM_lA8IQbfCM2bLz4pkBWgGgSavNmrZ_Y-7CkJcDv6g");

  var    pay_init = pay_initSpreadSheet.getSheetByName("Data");
  var    pay_suc = pay_sucSpreadSheet.getSheetByName("Data");
  var    pay_f = pay_fSpreadSheet.getSheetByName("Data"); 
  var    tckt_iss = tckt_issSpreadSheet.getSheetByName("Data");
  var    offline = offlineSpreadSheet.getSheetByName("Data");


  
  
  ///////  // ************ Payment Initiated
  
  
  var threads1 = GmailApp.search('is:unread from:me "PaymentStatus: In progress," ' );
  
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 
    var count = pay_init.getRange(1,1).getValue();
   for (var i=0; i<threads1.length; i++)
  {
    var messages = threads1[i].getMessages();
    var msg_ln = messages.length

    
      // var msg = messages[msg_ln-1].getBody();
      var sub = messages[0].getSubject();
      var dat = messages[0].getDate();
     var id = sub.substring(sub.indexOf("for")+4,sub.indexOf("-") );
      if(id!=""){
    pay_init.getRange(i+3+count,1).setValue(id); 
    pay_init.getRange(i+3+count,2).setValue(dat); 
       
      }  
  
     pay_init.getRange(i+3+count, 3).setValue(new Date())
    GmailApp.markThreadRead(threads1[i]);
  }
  
  /////////  // ************ Payment Successful
  
    var count = pay_suc.getRange(1,1).getValue();
  var threads1 = GmailApp.search( ' is:unread from:me "Payment Successful for" ' );
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 

   for (var i=0; i<threads1.length; i++)
  {
    var messages = threads1[i].getMessages();
    var msg_ln = messages.length
   
    
      // var msg = messages[msg_ln-1].getBody();
      var sub = messages[0].getSubject();
      var dat = messages[0].getDate();
   
    var id = sub.substring(sub.indexOf("for")+4,sub.indexOf("-") );
      if (id!=""){
    pay_suc.getRange(i+3+count,1).setValue(id); 
    pay_suc.getRange(i+3+count,2).setValue(dat); 
      
      } 
    
      pay_suc.getRange(i+3+count, 3).setValue(new Date())
    GmailApp.markThreadRead(threads1[i]);
  }
   
  ////// ************ Payment failed
  
    var count = pay_f.getRange(1,1).getValue();
  var threads1 = GmailApp.search('is:unread from:me "Payment Failed for"  ');
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 

   for (var i=0; i<threads1.length; i++)
  {
    var messages = threads1[i].getMessages();
    var msg_ln = messages.length

  
      // var msg = messages[msg_ln-1].getBody();
      var sub = messages[0].getSubject();
     var dat = messages[0].getDate();
     
     var id = sub.substring(sub.indexOf("for")+4,sub.indexOf("-") );
     
      if (id!=""){
        pay_f.getRange(i+3+count,1).setValue(id); 
       pay_f.getRange(i+3+count,2).setValue(dat);
        
     
     } 
    
    pay_f.getRange(i+3+count, 3).setValue(new Date())
    GmailApp.markThreadRead(threads1[i]);
  }
  
  
  
  
  //////// ************ Ticket Issued 
 

    var count = tckt_iss.getRange(1,1).getValue();
  var threads1 = GmailApp.search('is:unread from:me "CatchThatBus from" ' );
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 

   for (var i=0; i<threads1.length; i++)
  {
    var messages = threads1[i].getMessages();
    var msg_ln = messages.length

    if (msg_ln<2){
      // var msg = messages[msg_ln-1].getBody();
      var sub = messages[0].getSubject();
      var dat = messages[0].getDate();
     var body = messages[0].getBody();
    var id= body.substring(body.indexOf("Transaction No:")+151,
                            body.indexOf("Transaction No:")+158);
      if(id!=""){
    
    tckt_iss.getRange(i+3+count,1).setValue(id); 

    tckt_iss.getRange(i+3+count,2).setValue(dat); 
        
      }  
  
    }
    tckt_iss.getRange(i+3+count, 3).setValue(new Date())
    GmailApp.markThreadRead(threads1[i]);
  
  }
  
   
  // ************ Offline Booking
  
  
      var count = offline.getRange(1,1).getValue();
  var threads1 = GmailApp.search('is:unread from:me "Booking from" ' );
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 

  for (var i=0; i<threads1.length; i++)
  {
    var messages = threads1[i].getMessages();
    var msg_ln = messages.length

    if (msg_ln<2){
      // var msg = messages[msg_ln-1].getBody();
      var sub = messages[0].getSubject();
      var dat = messages[0].getDate();
     var body = messages[0].getBody();
    var id= body.substring(body.indexOf("for your reference.")-17,
                            body.indexOf("for your reference.")-10);
      if(id!=""){
    offline.getRange(i+3+count,1).setValue(id); 
         offline.getRange(i+3+count,2).setValue(dat); 
        offline.getRange(i+3+count, 3).setValue(new Date())
      }  
      
    
    
    }
   offline.getRange(i+3+count , 3).setValue(new Date())
    GmailApp.markThreadRead(threads1[i]);
       }
  
  
  
  //// Read Molpay Mails
  
  var threads1 = GmailApp.search('is:unread "molpay" ' );
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 

  for (var i=0; i<threads1.length; i++)
  {
    
       
    GmailApp.markThreadRead(threads1[i]);
       }
  
  //// Pre Booking 
  
  var threads1 = GmailApp.search('is:unread "Pre Booking Request "' );
  // var threads2 = GmailApp.search('after:2017/02/22 AND "Payment completed transaction id" AND is:unread');
 

  for (var i=0; i<threads1.length; i++)
  {
    
       
    GmailApp.markThreadRead(threads1[i]);
       }
  
}
