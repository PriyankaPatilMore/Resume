using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;



namespace Emailproject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        // GET api/values
        [HttpGet("{emailaddress}")]
        public ActionResult<IEnumerable<string>> Get(string emailaddress)
        {
            try{
                /******************* Write Your email below here    *********************************/
                string fromaddr = "more.piyapatil@gmail.com";
                /****************** Write Your password below here  ********************************/
                string password = "Piya@231078";
                MailMessage msg = new MailMessage();
                msg.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");
                msg.Subject = "Request for job vacancies in the field of Software Application Development - Priyanka More";
                msg.From = new MailAddress(fromaddr);
                //msg.Body = "Verfication code to register at Krunsave: "+otp.ToString()+"  Link: http://18.222.237.46/VerifyRegister.html";
                msg.Body = "Respected Sir,\r\n \r\nI am writing to ask if you have, or are likely to have any vacancies in the software development area. I have finished my Masters in Computer Science from the Asian Institute of Technology, Thailand in December 2018, and I am looking for full-time employment opportunities in Thailand. I would really like to work as a Software Developer and would be prepared to commit to any training that might be required. I have enclosed my Resume, and Transcript and I look forward to hearing from you. After finishing of my Masters study I return back to India(Home-Town) due to completion of my student visa. Therefore, If my profile is consider, I can be available for online interview via Skype/Line or any online service application at any time. I value your feedback and kindly waiting for the response.\r\n \r\nThank You for your time and consideration.\r\n\r\nYours Sincerely\r\nPriyanka More";
                var filename1 = @"C:\Users\karri\Desktop\Resume1\SuryaRaoKarriResume.pdf";
                var filename2 = @"C:\Users\karri\Desktop\Resume1\SuryaRaoKarriCoverLetter.pdf";
                //var filename3 = @"C:\Users\karri\Desktop\Resume1\SuryaRaoKarriOfficialTranscript.pdf";
                msg.Attachments.Add(new Attachment(filename1));
                msg.Attachments.Add(new Attachment(filename2));
                //msg.Attachments.Add(new Attachment(filename3));
                var allemailaddress = emailaddress.Split(",");
                foreach(var oneEmail in allemailaddress){
                    msg.To.Add(new MailAddress(oneEmail));
                }
                SmtpClient smtp = new SmtpClient();
                smtp.Host = "smtp.gmail.com";
                smtp.Port = 587;
                smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = true;
                NetworkCredential nc = new NetworkCredential(fromaddr,password);
                smtp.Credentials = nc;
                smtp.Send(msg);
                
                return new string[]{"send = "+emailaddress};
            }
            catch{
                return new string[]{"not send = "+emailaddress};
            }
            //return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet]
        public ActionResult<string> Get()
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
