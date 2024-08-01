#region NameSpaces
using System;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Configuration;
#endregion

namespace EmailPost
{
    class Program
    {
        static void Main(string[] args)
        {
            int sDia = DateTime.Now.Day;
            int sMes = DateTime.Now.Month;
            int sMesOffset = 0;
            if (sDia < 3)
            {
                sMesOffset = -1;
            }
            string sNmes = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(DateTime.Now.AddMonths(sMesOffset).ToString("MMMM"));
            string sHora = DateTime.Now.ToString("hh:mm tt");
            string informeCruzado = ConfigurationManager.AppSettings["informe_cz"];
            string sSubject = "";
            string sBody = "";
            string xlpath = ""; // Excel to upd filepath
            string gerenciaPath = "";
            string emailTO = "";
            string emailCC = "";
            int iPriority = 2;

            // POSTPAGO
            sSubject = "Reporte unificado de Calidad Postpago - " + sNmes + " (al " + sDia + "/" + sMes + " - " + sHora + ")";
            sBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\">  <head> <meta http-equiv=Content-Type content=\"text/html; charset=utf-8\"> <meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"> <style> <!-- /* Font Definitions */  @font-face { font-family: \"Cambria Math\"; panose-1: 2 4 5 3 5 4 6 3 2 4; }  @font-face { font-family: Calibri; panose-1: 2 15 5 2 2 2 4 3 2 4; } /* Style Definitions */  p.MsoNormal, li.MsoNormal, div.MsoNormal { margin: 0cm; margin-bottom: .0001pt; font-size: 12.0pt; font-family: \"Times New Roman\", \"serif\"; }  a:link, span.MsoHyperlink { mso-style-priority: 99; color: blue; text-decoration: underline; }  a:visited, span.MsoHyperlinkFollowed { mso-style-priority: 99; color: purple; text-decoration: underline; }  span.gmail-il { mso-style-name: gmail-il; }  span.EstiloCorreo18 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo19 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo20 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo21 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo22 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo23 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo24 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo25 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo26 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo27 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo28 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo29 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo30 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo31 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo32 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo33 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo34 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo35 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo36 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo37 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo38 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo39 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo40 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo41 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo42 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo43 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo44 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo45 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo46 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo47 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo48 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: windowtext; }  span.EstiloCorreo49 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo50 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo51 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo52 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo53 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo54 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo55 { mso-style-type: personal-reply; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  .MsoChpDefault { mso-style-type: export-only; font-size: 10.0pt; }  @page WordSection1 { size: 612.0pt 792.0pt; margin: 70.85pt 3.0cm 70.85pt 3.0cm; }  div.WordSection1 { page: WordSection1; }  --> </style> <!--[if gte mso 9]><xml> <o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" /> </xml><![endif]--> <!--[if gte mso 9]><xml> <o:shapelayout v:ext=\"edit\"> <o:idmap v:ext=\"edit\" data=\"1\" /> </o:shapelayout></xml><![endif]--> </head>  <body lang=ES-PE link=blue vlink=purple> <div class=WordSection1> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Buen día.<o:p></o:p></span></p> &nbsp;<p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Se adjunta Reporte de calidad Postpago, en el cual se unifican los siguientes reportes de calidad:<o:p></o:p></span></p> &nbsp;<p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Reporte de calidad Interna Ejecutivos de calidad<span style='color:#0070C0'>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span style='color:#1F497D'>Actualizado al " + sDia + "/" + sMes + ".</span> <o:p></o:p> </span> </p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Reporte de calidad Interna Evaluaciones supervisores:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style='color:#1F497D'>Actualizado al " + sDia + "/" + sMes + ".</span> <o:p></o:p> </span> </p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Reporte de monitoreo cruzado:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style='color:#1F497D'>" + informeCruzado + ".</span> <o:p></o:p> </span> </p> &nbsp;<p class=MsoNormal></p> &nbsp;<p class=MsoNormal><b><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\";color:#215968'>Adicionalmente el reporte se encuentra ubicado en la siguiente ruta:</span></b><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'><o:p></o:p></span></p> <p class=MsoNormal><a href=\"file:///\\\\192.168.150.5\\GerenciaCX\\1.%20CALIDAD\\GERENCIA%200\\1.%20POSTPAGO\\01.%20INFORME%20UNIFICADO\\2020\\04.Abril\\\"><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>\\\\192.168.150.5\\GerenciaCX\\1. CALIDAD\\GERENCIA 0\\01. POSTPAGO\\1. INFORME UNIFICADO\\2020\\04.Abril</span></a><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\";color:#1F497D'><o:p></o:p></span></p> &nbsp;<p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Saludos,<o:p></o:p></span></p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>protected<o:p></o:p></span></p> </div> </body>  </html>";
            xlpath = @"C:\01.Reportes Analista Calidad\98.REPORTES\Unificados\POST\202004_Reporte_unificado_calidad_Postpago.xlsm"; // Excel to upd filepath
            gerenciaPath = @"Z:\1. CALIDAD\GERENCIA 0\1. POSTPAGO\01. INFORME UNIFICADO\2020\04.Abril\202004_Reporte_unificado_calidad_Postpago.xlsm";
            emailTO = ConfigurationManager.AppSettings["toPOST"];
            emailCC = ConfigurationManager.AppSettings["ccPOST"];
            //emailCC = "";
            Console.WriteLine("Postpago: Actualizando Excel...");

            Upd_excl(xlpath);

            Console.WriteLine("Postpago: Copiando Excel a Gerencia...");

            System.IO.File.Copy(xlpath, gerenciaPath, true);

            Console.WriteLine("Postpago: Enviando correo...");
            if (SendMail(sSubject, sBody, iPriority, xlpath, emailTO, emailCC))
            {
                Console.WriteLine("Postpago: Todo OK!");
            }
            else
            {
                Console.WriteLine("Postpago: Todo MAL!");
            }

            // HFC
            sSubject = "Reporte unificado de Calidad HFC - " + sNmes + " (al " + sDia + "/" + sMes + " - " + sHora + ")";
            sBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\">  <head> <meta http-equiv=Content-Type content=\"text/html; charset=utf-8\"> <meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"> <style> <!-- /* Font Definitions */  @font-face { font-family: \"Cambria Math\"; panose-1: 2 4 5 3 5 4 6 3 2 4; }  @font-face { font-family: Calibri; panose-1: 2 15 5 2 2 2 4 3 2 4; } /* Style Definitions */  p.MsoNormal, li.MsoNormal, div.MsoNormal { margin: 0cm; margin-bottom: .0001pt; font-size: 12.0pt; font-family: \"Times New Roman\", \"serif\"; }  a:link, span.MsoHyperlink { mso-style-priority: 99; color: blue; text-decoration: underline; }  a:visited, span.MsoHyperlinkFollowed { mso-style-priority: 99; color: purple; text-decoration: underline; }  span.EstiloCorreo17 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo18 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo19 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo20 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo21 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo22 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo23 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo24 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo25 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo26 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo27 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo28 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo29 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo30 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo31 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo32 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo33 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo34 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo35 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo36 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo37 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo38 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo39 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo40 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo41 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo42 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo43 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo44 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo45 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo46 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo47 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: windowtext; }  span.EstiloCorreo48 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo49 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo50 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo51 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo52 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo53 { mso-style-type: personal; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  span.EstiloCorreo54 { mso-style-type: personal-reply; font-family: \"Calibri\", \"sans-serif\"; color: #1F497D; }  .MsoChpDefault { mso-style-type: export-only; font-size: 10.0pt; }  @page WordSection1 { size: 612.0pt 792.0pt; margin: 70.85pt 3.0cm 70.85pt 3.0cm; }  div.WordSection1 { page: WordSection1; }  --> </style> <!--[if gte mso 9]><xml> <o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" /> </xml><![endif]--> <!--[if gte mso 9]><xml> <o:shapelayout v:ext=\"edit\"> <o:idmap v:ext=\"edit\" data=\"1\" /> </o:shapelayout></xml><![endif]--> </head>  <body lang=ES-PE link=blue vlink=purple> <div class=WordSection1> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Buen día.<o:p></o:p></span></p> &nbsp;<p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Se envía el reporte de calidad HFC en el cual se unifican los siguientes reportes de calidad:<o:p></o:p></span></p> &nbsp;<p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Reporte de calidad Interna Ejecutivos de calidad: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <span style='color:#1F497D'>Actualizado al " + sDia + "/" + sMes + ".</span> <o:p></o:p> </span> </p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Reporte de calidad Interna Evaluaciones supervisores: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <span style='color:#1F497D'>Actualizado al " + sDia + "/" + sMes + ".</span> <o:p></o:p> </span> </p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Reporte de monitoreo Cruzado <span style='color:#1F497D'>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " + informeCruzado + ".</span> <o:p></o:p> </span> </p> &nbsp;<p class=MsoNormal><b><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\";color:#1F497D'>Adicionalmente el reporte se encuentra ubicado en la siguiente ruta:<o:p></o:p></span></b></p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\";color:#1F497D'><a href=\"file:///\\\\192.168.150.5\\GerenciaCX\\1.%20CALIDAD\\GERENCIA%200\\2.%20HFC\\01.%20INFORME%20UNIFICADO\\2020\\04.Abril\\\">\\\\192.168.150.5\\GerenciaCX\\1. CALIDAD\\GERENCIA 0\\2. HFC\\01. INFORME UNIFICADO\\2019\\04.Abril</a><o:p></o:p></span></p> &nbsp;<p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>Saludos,<o:p></o:p></span></p> <p class=MsoNormal><span style='font-size:11.0pt;font-family:\"Calibri\",\"sans-serif\"'>protected<o:p></o:p></span></p> </div> </body>  </html>";
            xlpath = @"C:\01.Reportes Analista Calidad\98.REPORTES\Unificados\HFC\202004_Reporte_unificado_calidad_HFC.xlsm"; // Excel to upd filepath
            gerenciaPath = @"Z:\1. CALIDAD\GERENCIA 0\2. HFC\01. INFORME UNIFICADO\2020\04.Abril\202004_Reporte_unificado_calidad_HFC.xlsm";
            emailTO = ConfigurationManager.AppSettings["toHFC"];
            emailCC = ConfigurationManager.AppSettings["ccHFC"];

            Console.WriteLine("HFC: Actualizando Excel...");
            Upd_excl(xlpath);

            Console.WriteLine("HFC: Copiando Excel a Gerencia...");
            System.IO.File.Copy(xlpath, gerenciaPath, true);

            Console.WriteLine("HFC: Enviando correo...");
            if (SendMail(sSubject, sBody, iPriority, xlpath, emailTO, emailCC))
            {
                Console.WriteLine("HFC: Todo OK!");
            }
            else
            {
                Console.WriteLine("HFC: Todo MAL!");
            }
        }

        static bool SendMail(string sSubject, string sMessage, int iPriority, string attach, string emailTo, string emailCC)
        {
            try
            {
                string sEmailServer = "smtp.gmail.com";
                string sEmailPort = "587";
                string sEmailUser = "protected";
                string sEmailPass = "protected";
                string sEmailSendTo = emailTo;
                string sEmailSendCC = emailCC;
                string sEmailSendFrom = "protected";
                string sEmailSendFromName = "protected";

                SmtpClient smtpClient = new SmtpClient();
                MailMessage message = new MailMessage();

                MailAddress fromAddress = new MailAddress(sEmailSendFrom, sEmailSendFromName);

                //You can have multiple emails separated by ;
                string[] sEmailTo = Regex.Split(sEmailSendTo, ";");
                string[] sEmailCC = Regex.Split(sEmailSendCC, ";");
                int sEmailServerSMTP = int.Parse(sEmailPort);

                smtpClient.Host = sEmailServer;
                smtpClient.Port = sEmailServerSMTP;
                smtpClient.EnableSsl = true;
                //smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                System.Net.NetworkCredential myCredentials = new System.Net.NetworkCredential(sEmailUser, sEmailPass);
                smtpClient.Credentials = myCredentials;

                message.From = fromAddress;

                if (sEmailTo != null)
                {
                    for (int i = 0; i < sEmailTo.Length; i++)
                    {
                        if (sEmailTo[i] != null && sEmailTo[i] != "")
                        {
                            message.To.Add(sEmailTo[i]);
                        }

                    }
                }

                if (sEmailCC != null)
                {
                    for (int i = 0; i < sEmailCC.Length; i++)
                    {
                        if (sEmailCC[i] != null && sEmailCC[i] != "")
                        {
                            message.CC.Add(sEmailCC[i]);
                        }
                    }
                }

                switch (iPriority)
                {
                    case 1:
                        message.Priority = MailPriority.High;
                        break;
                    case 3:
                        message.Priority = MailPriority.Low;
                        break;
                    default:
                        message.Priority = MailPriority.Normal;
                        break;
                }

                //You can enable this for Attachements.  SingleFile is a string variable for the file path.
                //foreach (string SingleFile in myFiles)
                //{
                //    Attachment myAttachment = new Attachment(SingleFile);
                //    message.Attachments.Add(myAttachment);
                //}

                if (attach != null)
                {
                    Attachment myAttachment = new Attachment(attach);
                    // Attachment myAttachment2 = new Attachment(Dts.Variables["User::vReincidentesPost_fullpath"].Value.ToString());
                    message.Attachments.Add(myAttachment);
                    // message.Attachments.Add(myAttachment2);
                }


                message.Subject = sSubject;
                message.IsBodyHtml = true;
                message.Body = sMessage;

                smtpClient.Send(message);

                return true;
            }
            catch (Exception ex)
            {
                //Debug.WriteLine(ex);
                Console.WriteLine(ex);
                return false;
            }
        }

        static private void ReleaseCom(Object o)
        {
            try
            {
                if (o != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
                }
            }
            catch (Exception ex)
            {
                //Dts.Events.FireError(9999, "Error Excel Sheet Deletion", ex.Message, "", 1);
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (o != null)
                {
                    o = null;
                }
            }
        }

        static private void Upd_excl(string filename)
        {
            Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = xlapp.Workbooks.Open(filename);
            wb.RefreshAll();
            xlapp.CalculateUntilAsyncQueriesDone();
            wb.Close(true);
            ReleaseCom(wb);
            xlapp.Quit();
            ReleaseCom(xlapp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
