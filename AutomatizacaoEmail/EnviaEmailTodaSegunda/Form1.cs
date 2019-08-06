using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using log4net;
using System.Configuration;
using System.Net.Mail;
using System.Web;
using System.Globalization;
using OfficeOpenXml;
using System.Drawing;


namespace EnviaEmailTodaSegunda

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string HoraAtual = DateTime.Now.Hour.ToString();
            string minutoAtual = DateTime.Now.Minute.ToString();
            string segundoAtual = DateTime.Now.Second.ToString();
            string horacompleta = HoraAtual + ":" + minutoAtual + ":" + segundoAtual;
         
            LogHelper.WriteDebugLog("Aplicação iniciada.");
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            timerSendMail.Enabled = true;

            int i;
            DateTime dataMesAnterior = DateTime.Now;
            string DataDoRelatorio = "";
            dataMesAnterior = dataMesAnterior.AddMonths(-13);

            comboBox1.Items.Clear();
            for (i = 0; i <= 11; i++)
            {
                dataMesAnterior = dataMesAnterior.AddMonths(1);
                DataDoRelatorio = dataMesAnterior.ToString("MM/yyyy");
                comboBox1.Items.Add(DataDoRelatorio);
            }

            #region Conexão com a base de dados         
            try
            {
                string Stringdeconexao = Properties.Settings.Default.cone;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("Recebendo o parâmetro do App.Config STRING DE CONEXÃO: " + Stringdeconexao);

                using (OleDbConnection connection = new OleDbConnection(Stringdeconexao))
                {
                    connection.Open();
                    LogHelper.WriteDebugLog("Tentando conectar a base de dados.");                 
                    if (connection.State == ConnectionState.Open)
                    {
                        label1.Text = "A aplicação está conectada com a base de dados";
                        LogHelper.WriteDebugLog("Conexão concluída com sucesso.");
                    }
                    else
                    {
                        label1.Text = "A aplicação não está conectada com a base de dados.";
                        LogHelper.WriteDebugLog("conexão falhou.");
                    }
                }

            }
            catch (OleDbException x)
            {
                MessageBox.Show("Aconteceu um erro de conexão: "+ x);
                LogHelper.WriteDebugLog("ERRO: " + x);
            }
            #endregion
        }
        private void label1_Click(object sender, EventArgs e)
        {
                        
        }

        #region Função para Envio de Email

        public bool PSendEmail(string fpTo, string fpFrom, string fpSubject, string fpMailServerIP, string fpMsg, string fpAttachment, bool HtmlUse = false, string fpSMTPUsername = "", string fpSMTPPWD = "")
        {
            MailMessage mm = new MailMessage();
            SmtpClient smtp = new SmtpClient(fpMailServerIP, 25);
            AlternateView value;
            var ContentType = new System.Net.Mime.ContentType("text/html");
            string[] lvToFld;
            int x = 0;
            bool lvret = false;

            try
            {
                if (HtmlUse)
                {
                    value = AlternateView.CreateAlternateViewFromString(fpMsg, ContentType);
                    mm.AlternateViews.Add(value);
                }
                if (fpSMTPUsername != "")
                {
                    {
                        var withBlock = smtp;
                        withBlock.EnableSsl = false;
                        withBlock.Credentials = new System.Net.NetworkCredential(fpSMTPUsername, fpSMTPPWD);
                        withBlock.Timeout = 20000;       // I add this extra line
                    }
                }

                mm.From = new MailAddress(fpFrom, "Suporte MidiaVox");
                // Verificar se existe mais de 1 destinatário
                lvToFld = fpTo.Split(';');

                //if (Information.UBound(lvToFld) == 0)  //Information.UBound
                if (lvToFld.Length == 0)
                    mm.To.Add(fpTo);
                else
                    for (x = 0; x < lvToFld.Length; x++)
                        mm.To.Add(lvToFld[x].ToString());
                if (fpSubject == "")
                    fpSubject = "No - Subject";
                mm.Subject = fpSubject;
                mm.Body = fpMsg;
                // smtp.Host = fpMailServerIP

                if (fpAttachment != "")
                {
                    if (File.Exists(fpAttachment))
                    {
                        Attachment AttachmentData = new Attachment(fpAttachment);
                        mm.Attachments.Add(AttachmentData);
                        smtp.Send(mm);
                        LogHelper.WriteDebugLog("psendMail: Email enviado para :" + fpTo + " com attacment=" + fpAttachment);
                        lvret = true;
                        smtp.Dispose();
                        AttachmentData = null;
                    }
                    else
                        LogHelper.WriteDebugLog("pSendMail: Attachment para o email" + fpTo + " nao encontrado:" + fpAttachment);
                }
                else
                {
                    smtp.Send(mm);
                    LogHelper.WriteDebugLog("psendMail: Email enviado para :" + fpTo + " - S/Attacments");
                    lvret = true;
                    smtp.Dispose();
                }
                mm = null;
                smtp = null;

                // return lvret;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogHelper.WriteDebugLog("PSendEmail Error: " + ex.Message);
            }

            return lvret;
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)

        {
        #region Botão Envia Email com o número de chamados abertos por representantes
            EnviaEmailQtd();
            #endregion
        }

        private void button3_Click(object sender, EventArgs e)
        {
        #region Botão Envia Email para todos os representantes com os seus chamados em específico
            EnviaEmailSemanal();
            #endregion
        }

        public void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        #region função Envia Email com a quantidade de chamados
        public void EnviaEmailQtd()
        {
            LogHelper.WriteDebugLog("Executando Função EnviaEmailQtd.");
            string Stringdeconexao = Properties.Settings.Default.cone;
            Properties.Settings.Default.Save();
            string ServerIp = Properties.Settings.Default.svip;
            Properties.Settings.Default.Save();
            string usuario = Properties.Settings.Default.usuario;
            Properties.Settings.Default.Save();
            string senha = Properties.Settings.Default.senha;
            Properties.Settings.Default.Save();
            LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
            LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
            LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: "+ usuario);
            LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);

            DataTable dt = new DataTable();
            LogHelper.WriteDebugLog("Criando DataTable: "+dt);
            OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
            LogHelper.WriteDebugLog("abrindo conexão com a base de dados");
            OleDbCommand comando = new OleDbCommand("SELECT U.FIRSTNAME AS Nome, COUNT(*) AS Qtd FROM PROBLEMS, TBLUSERS AS U WHERE PROBLEMS.CLOSE_DATE IS NULL AND PROBLEMS.REP = U.SID GROUP BY U.FIRSTNAME ORDER BY FIRSTNAME ASC", conexao);


            try
            {

                conexao.Open();
                LogHelper.WriteDebugLog("conexão aberta.");
                comando.Connection = conexao;
                LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);

                OleDbDataAdapter data = new OleDbDataAdapter(comando);
                LogHelper.WriteDebugLog("QUERRY EXECUTADA, INFORMAÇÕES RESGATADAS DA BASE DE DADOS.");
                data.Fill(dt);
                LogHelper.WriteDebugLog("preenchendo data table: " + dt);
            } catch (Exception o)
            {
                LogHelper.WriteDebugLog("ERRO " + o);
                MessageBox.Show("ERRO AO CONECTAR: " + o.Message);
            }
            try {
                string mensagem = "<table border=7><tr><th>Nome</th><th>Chamados</th></tr>";
                string nome;
                string data2 = DateTime.Now.ToString(new CultureInfo("pt-BR", false).DateTimeFormat.ShortDatePattern);
                string qtd;
                string subject = "Quantidade de Chamados abertos - " + data2;
                string iniciomsg = "<html><head></head><body>Ol&aacute;, segue abaixo a lista de chamados abertos por representante. <br/><br/>";
                string fimmsg = "</tbody></table><br/> Atenciosamente, <br/>Suporte MidiaVox.</body></html>";

                List<string> lst = new List<string>();
                foreach (DataRow r in dt.Rows)
                {
                    LogHelper.WriteDebugLog("Percorrendo datatable" + dt);

                    nome = r["Nome"].ToString();
                        qtd = r["Qtd"].ToString();
                        mensagem += "<tr><td>" + HttpUtility.HtmlEncode(nome) + " " + "</td><td>" + qtd + "</td></tr>";
                    LogHelper.WriteDebugLog("obtendo acesso a os dados:" + nome + " e" + qtd);
                    lst.Add(r["Nome"].ToString() + ": " + r["Qtd"].ToString());
                }
               
                listBox1.DataSource = lst;
                LogHelper.WriteDebugLog("adicionando as informações no listbox");

                conexao.Close(); LogHelper.WriteDebugLog("conexão com a base de dados fechada");

                PSendEmail("andre@midiavox.com.br", "suporte@midiavox.com.br", subject, ServerIp, iniciomsg + mensagem + fimmsg, "", true, usuario, senha);
                LogHelper.WriteDebugLog("Email enviado para: andre@midiavox.com.br com os dados resgatados da base de dados." );
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Falha ao enviar email: " + ex.Message);
                LogHelper.WriteDebugLog("ERRO: "+ex);
            }

        }
        #endregion
        #region função Envia Email com os Chamados Desatualizados
        public void AtualizacaoChamado()
        {
            
            try
            {
                LogHelper.WriteDebugLog("Função AtualizacaoChamados selecionada");
                string Stringdeconexao = Properties.Settings.Default.cone;
                Properties.Settings.Default.Save();
                string ServerIp = Properties.Settings.Default.svip;
                Properties.Settings.Default.Save();
                string usuario = Properties.Settings.Default.usuario;
                Properties.Settings.Default.Save();
                string senha = Properties.Settings.Default.senha;
                Properties.Settings.Default.Save();
                string dias = Properties.Settings.Default.diasAtualizacaoChamados;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
                LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
                LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: " + usuario);
                LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);
                LogHelper.WriteDebugLog("recebendo diasas do App.Config: " + dias);


                DataTable dt = new DataTable();
                LogHelper.WriteDebugLog("criando datatable " + dt);

                ////Cria a conexão com a Base de Dados
                OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
                LogHelper.WriteDebugLog("criando conexão com a base de dados");
                OleDbCommand comando = new OleDbCommand("SELECT P.ID, U.FIRSTNAME as Nome, U.EMAIL1 as Email FROM PROBLEMS AS P, TBLUSERS AS U WHERE P.CLOSE_DATE IS NULL AND P.REP = U.SID AND P.ID IN(SELECT ID FROM tblNotes as N GROUP BY N.ID HAVING MAX(addDate) <= DATEADD(dd, -" + dias + ", GETDATE())) ORDER BY ID", conexao);
                LogHelper.WriteDebugLog("SELECIONANDO QUERRY " + comando);

                //Tenta abrir a conexao           
                conexao.Open();
                LogHelper.WriteDebugLog("conexão com a base de dados aberta");
                ////Aponta ao MySqlCommand que esta vindo por parametro, qual será a conexão a ser utilizada
                comando.Connection = conexao;
                ////Cria o DataAdapter que executara o comando SQL, e mostra a ela qual comando executar
                OleDbDataAdapter data = new OleDbDataAdapter(comando);
                LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);
                ////Preenche o DataTable com o Retorno do Select
                data.Fill(dt);
                LogHelper.WriteDebugLog("preenchendo datatable " + dt);


                string iniciomsg = "";
                string fimmsg = "";
                string assunto = "";
                string mensagem = "";
                string nome;
                string emai;
                string id;
                string data3 = DateTime.Now.ToString(new CultureInfo("pt-BR", false).DateTimeFormat.ShortDatePattern);

                List<string> lst = new List<string>();


                foreach (DataRow r in dt.Rows)
                {
                    id = r["ID"].ToString();
                    nome = r["Nome"].ToString();
                    emai = r["Email"].ToString();
                    LogHelper.WriteDebugLog("percorrendo datatable " + dt);
                    LogHelper.WriteDebugLog("recuperando informações da base de dados: " + id + " " + nome + " " + emai);


                    lst.Add(r["ID"].ToString() + ": " + r["Nome"].ToString() + ": " + r["Email"].ToString());
                    assunto = "Atualização chamado: " + id;
                    iniciomsg = "Ol&aacute; " + HttpUtility.HtmlEncode(nome) + "<br/><br/>";
                    mensagem = "Verifiquei que o chamado com o ID: " + "<a href=\"http://suporte.midiavox.com.br/rep_details.asp?id=" + id + "\">" + id + " " + "</a>" + " est&aacute; sem atualiza&ccedil;&atilde;o a mais de " + dias + " dias.<br/> &Eacute; muito importante que atualizemos as informa&ccedil;&otilde;es no site.<br/>";
                    fimmsg = "<br/<br/>Pode me ajudar a fazer uma r&aacute;pida atualiza&ccedil;&atilde;o? Muito obrigado.<br/><br/>Atenciosamente,<br/>Suporte MidiaVox.";


                    PSendEmail(emai, "suporte@midiavox.com.br", assunto, ServerIp, iniciomsg + mensagem + fimmsg, "", true, usuario, senha);
                    LogHelper.WriteDebugLog("Enviando Email para: " + emai);
                }

                listBox1.DataSource = lst;
                LogHelper.WriteDebugLog("adicionando elementos: " + lst + " ao listbox.");

                conexao.Close();
                LogHelper.WriteDebugLog("conexão fechada.");
            }
            catch (OleDbException a)
            {
                LogHelper.WriteDebugLog("ERRO: " + a);
                MessageBox.Show("Falha ao enviar email: " + a.Message);
            }
            
        }
        #endregion
        #region função Envia Email com os chamados por representante
        public void EnviaEmailSemanal()
        {
            LogHelper.WriteDebugLog("Executando função EnviaEmailSemanal");
            string Stringdeconexao = Properties.Settings.Default.cone;
            Properties.Settings.Default.Save();
            string ServerIp = Properties.Settings.Default.svip;
            Properties.Settings.Default.Save();
            string usuario = Properties.Settings.Default.usuario;
            Properties.Settings.Default.Save();
            string senha = Properties.Settings.Default.senha;
            Properties.Settings.Default.Save();
            LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
            LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
            LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: " + usuario);
            LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);

        
            DataTable dt = new DataTable();
            LogHelper.WriteDebugLog("criando datatable " + dt);
            
            OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
            LogHelper.WriteDebugLog("criando conexão com a base de dados");
            OleDbCommand comando = new OleDbCommand("SELECT P.ID as ID, U.FIRSTNAME as Nome, U.EMAIL1 as Email, P.TITLE as Titulo FROM PROBLEMS AS P, TBLUSERS AS U WHERE P.CLOSE_DATE IS NULL AND P.REP = U.SID ORDER BY EMAIL1 ASC", conexao);
            LogHelper.WriteDebugLog("SELECIONANDO QUERRY " + comando);


            try {   
                conexao.Open();
                LogHelper.WriteDebugLog("conexão com a base de dados aberta");

                comando.Connection = conexao;

                OleDbDataAdapter data = new OleDbDataAdapter(comando);
                LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);

                data.Fill(dt);
                LogHelper.WriteDebugLog("preenchendo datatable " + dt);

            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
                LogHelper.WriteDebugLog("ERRO AO CONECTAR: " + ex);
            }
            try
            {
                string iniciomsg = "";
                string fimmsg = "";
                string assunto = "";
                string mensagem = "<table border=7><tr><th>ID</th><th>Titulo</th></tr>";
                string nome;
                string emai;
                string til;
                string id;
                string emailSalvo = "";
                string mensageSalva = "";
                string inicio = ", temos a responsabilidade de mensalmente enviar o relatório para a Avaya com os chamados que estão abertos e os chamados fechados no mês. Verifiquei que sobre sua responsabilidade estão o(s) chamado(s) abaixo: ";
                string fim = "É importante que qualquer movimentação para resolver o problema do chamado o site seja atualizado. ";
                string fim2 = "Caso o problema já tenha sido resolvido, por favor, informar a solução, quantidade de horas e fechar o chamado.";
                string nomeSalvo = "";
                string data3 = DateTime.Now.ToString(new CultureInfo("pt-BR", false).DateTimeFormat.ShortDatePattern);
                List<string> lst = new List<string>();


                foreach (DataRow r in dt.Rows)
                {
                    id = r["ID"].ToString();
                    nome = r["Nome"].ToString();
                    emai = r["Email"].ToString();
                    til = r["Titulo"].ToString();
                    lst.Add(r["ID"].ToString() + ": " + r["Nome"].ToString() + ": " + r["Email"].ToString() + ": " + r["Titulo"].ToString());
                    fimmsg = "</tbody></table><br/>" + HttpUtility.HtmlEncode(fim) + "<br/>" + HttpUtility.HtmlEncode(fim2) + "<br/><br/>" + "Atenciosamente, <br/>Suporte MidiaVox.</body></html>";
                    LogHelper.WriteDebugLog("percorrendo datatable " + dt);
                    LogHelper.WriteDebugLog("recuperando informações da base de dados: " + id + " " + nome + " " + emai + " " + til);
                    if (emailSalvo == emai || emailSalvo == "")
                    {
                        assunto = "Chamados de " + nomeSalvo + " - " + data3;
                        iniciomsg = "<html><head></head><body>Ol&aacute; " + HttpUtility.HtmlEncode(nomeSalvo) + HttpUtility.HtmlEncode(inicio) + "<br/><br/>";
                        mensagem += "<tr><td><a href=\"http://suporte.midiavox.com.br/rep_details.asp?id=" + id + "\">" + id + " " + "</a></td><td>";
                        mensagem += HttpUtility.HtmlEncode(til) + "</td></tr>";
                    }
                    else
                    {

                        PSendEmail(emailSalvo, "suporte@midiavox.com.br", assunto, ServerIp, iniciomsg + mensageSalva + fimmsg, "", true, usuario, senha);
                        LogHelper.WriteDebugLog("Enviando Email para: " + emailSalvo);

                        mensagem = "<table border=7><tr><th>ID</th><th>Titulo</th></tr><tr><td><a href=\"http://suporte.midiavox.com.br/rep_details.asp?id=" + id + "\">" + id + " " + "</a></td><td>";
                        mensagem += HttpUtility.HtmlEncode(til) + "</td></tr>";


                        mensageSalva = "";
                        emailSalvo = "";
                        assunto = "Chamados de " + nomeSalvo + " - " + data3;
                        iniciomsg = "<html><head></head><body>Ol&aacute; " + HttpUtility.HtmlEncode(nomeSalvo) + HttpUtility.HtmlEncode(inicio) + ".<br/><br/>";

                    }
                    emailSalvo = emai;
                    nomeSalvo = nome;
                    mensageSalva = mensagem;
                    assunto = "Chamados de " + nomeSalvo + " - " + data3;
                    iniciomsg = "<html><head></head><body>Ol&aacute " + HttpUtility.HtmlEncode(nomeSalvo) + HttpUtility.HtmlEncode(inicio) + "<br/><br/>";

                }
                if (emailSalvo != "")
                {
                    PSendEmail(emailSalvo, "suporte@midiavox.com.br", assunto, ServerIp, iniciomsg + mensageSalva + fimmsg, "", true, usuario, senha);
                    LogHelper.WriteDebugLog("Enviando Email para: " + emailSalvo);
                    emailSalvo = "";
                    mensagem = "";
                    nomeSalvo = "";
                    iniciomsg = "";
                }

                listBox1.DataSource = lst;
                LogHelper.WriteDebugLog("adicionando elementos: " + lst + " ao listbox.");

                conexao.Close();
                LogHelper.WriteDebugLog("conexão fechada.");

            }
            catch (Exception z)
            {
                LogHelper.WriteDebugLog("ERRO:" + z);
                MessageBox.Show("Falha ao enviar email: " + z.Message);
            }
        }
        #endregion
        #region função Envia Email com o relatório mensal
        public void EnviaRelatorioMensal(DateTime dataNow)
        {
          
            try
            {
                Document doc = new Document(PageSize.A4);//criando e estipulando o tipo da folha usada
                doc.SetMargins(60, 60, 60, 60);//estibulando o espaçamento das margens que queremos
                doc.AddCreationDate();//adicionando as configuracoes

                //caminho onde sera criado o pdf + nome desejado
                //OBS: o nome sempre deve ser terminado com .pdf
                //DateTime dataMesAnterior = DateTime.Now; By André
                DateTime dataMesAnterior = dataNow;
                dataMesAnterior = dataMesAnterior.AddMonths(-1);

                string MesDoRelatorio = dataMesAnterior.ToString("MMMM - yyyy");
                String pathApp = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
                string caminho = pathApp + "Relatorios\\Relatório de Atendimento Suporte MidiaVox - Avaya - " + MesDoRelatorio + ".pdf";

                //byte[] bytes = Encoding.Default.GetBytes(caminho);
                //caminho = Encoding.UTF8.GetString(bytes);

                LogHelper.WriteDebugLog("Criando documento " + doc + " equivalente ao mês: " + MesDoRelatorio);
                LogHelper.WriteDebugLog("Adicionando documento ao caminho: " + caminho);

                //criando o arquivo pdf embranco, passando como parametro a variavel                
                //doc criada acima e a variavel caminho 
                //tambem criada acima.
                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(caminho, FileMode.Create));

                doc.Open();
                LogHelper.WriteDebugLog("Abrindo documento: " + doc);

                //Cores e Fontes
                BaseColor AzulMeiaNoite = new BaseColor(25, 25, 112);
                iTextSharp.text.Font fonte1 = FontFactory.GetFont("Calibri", 36, AzulMeiaNoite);
                LogHelper.WriteDebugLog("criando cor e fonte: " + AzulMeiaNoite + " " + fonte1);

                BaseColor RoyalBlue = new BaseColor(65, 105, 225);
                iTextSharp.text.Font fonte2 = FontFactory.GetFont("Calibri", 20, RoyalBlue);
                LogHelper.WriteDebugLog("criando cor e fonte: " + RoyalBlue + " " + fonte2);

                BaseColor CinzaEscuro = new BaseColor(105, 105, 105);
                iTextSharp.text.Font fonte3 = FontFactory.GetFont("Calibri", 16, CinzaEscuro);
                LogHelper.WriteDebugLog("criando cor e fonte: " + CinzaEscuro + " " + fonte3);

                iTextSharp.text.Font fonte4 = FontFactory.GetFont("Calibri", 11);
                LogHelper.WriteDebugLog("criando fonte: " + fonte4);

                BaseColor VerdeOliva = new BaseColor(107, 142, 35);
                iTextSharp.text.Font fonte5 = FontFactory.GetFont("Calibri", 14, VerdeOliva);
                iTextSharp.text.Font fonte51 = FontFactory.GetFont("Calibri", 28, VerdeOliva);
                LogHelper.WriteDebugLog("criando cor e fonte: " + VerdeOliva + " " + fonte5);
                LogHelper.WriteDebugLog("criando fonte: " + fonte51);


                BaseColor AzulLink = new BaseColor(0, 0, 255);
                iTextSharp.text.Font fonte6 = FontFactory.GetFont("Calibri", 14, AzulLink);
                LogHelper.WriteDebugLog("criando cor e fonte: " + AzulLink + " " + fonte6);

                BaseColor Preto = new BaseColor(0, 0, 0);
                iTextSharp.text.Font fonte7 = FontFactory.GetFont("Liberation Serif", 14, Preto);
                LogHelper.WriteDebugLog("criando cor e fonte: " + Preto + " " + fonte7);

                BaseColor Branco = new BaseColor(255, 255, 255);
                iTextSharp.text.Font fonte8 = FontFactory.GetFont("Calibri", 8, Branco);
                LogHelper.WriteDebugLog("criando cor e fonte: " + Branco + " " + fonte8);

                iTextSharp.text.Font fonte9 = FontFactory.GetFont("Calibri", 7, Preto);
                LogHelper.WriteDebugLog("criando fonte: " + fonte9);
                //

                string Imagem = pathApp + "Imagens\\logo MIDIAVOX.jpg";
                iTextSharp.text.Image ImagemPdf = iTextSharp.text.Image.GetInstance(Imagem);
                doc.Add(ImagemPdf);
                LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf);

                //Parágrafos
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha

                Paragraph p1 = new Paragraph(string.Format("Relatório de Atendimento"), fonte1);
                doc.Add(p1);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p1);

                Paragraph p2 = new Paragraph(string.Format("Cliente: AVAYA BRASIL"), fonte2);
                doc.Add(p2);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p2);

                Paragraph p3 = new Paragraph(string.Format("Gerência de Suporte - MidiaVox"), fonte3);
                doc.Add(p3);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p3);


                string dataReferencia = dataMesAnterior.ToString("MM-yyyy");

                Paragraph p4 = new Paragraph(string.Format("Referência: " + dataReferencia), fonte3);
                doc.Add(p4);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p4);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha

                iTextSharp.text.Image ImagemPdf2 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\onda.png");
                doc.Add(ImagemPdf2);
                LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf2);

                doc.Add(new Paragraph(" ")); //Quebra de linha

                iTextSharp.text.Image ImagemPdf3 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\Minilogo.png");
                Paragraph p5 = new Paragraph(string.Format("Relatório Atendimento Suporte MidiaVox – Cliente AVAYA"), fonte4);
                p5.Alignment = Element.ALIGN_JUSTIFIED;
                doc.Add(ImagemPdf3);
                LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf3);
                doc.Add(p5);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p5);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha

                Paragraph pSum = new Paragraph(string.Format("                      SUMÁRIO"), fonte51);
                doc.Add(pSum);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pSum);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                Paragraph pSum1 = new Paragraph(string.Format("            Apresentação........................................................................03"), fonte5);
                doc.Add(pSum1);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pSum1);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                Paragraph pSum2 = new Paragraph(string.Format("            Detalhes dos chamados........................................................04"), fonte5);
                doc.Add(pSum2);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pSum2);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha

                iTextSharp.text.Image ImagemPdf4 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\Minilogo.png");
                Paragraph pTextoImagem = new Paragraph(string.Format("Relatório Atendimento Suporte MidiaVox – Cliente AVAYA"), fonte4);
                pTextoImagem.Alignment = Element.ALIGN_JUSTIFIED;
                doc.Add(ImagemPdf4);
                LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf4);
                doc.Add(pTextoImagem);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pTextoImagem);

                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                Paragraph p6 = new Paragraph(string.Format("1. Apresentação "), fonte5);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(p6);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p6);
                doc.Add(new Paragraph(" ")); //Quebra de linha

                string referencia1 = dataMesAnterior.ToString("MMMM");
                string referencia2 = dataMesAnterior.ToString("yyyy");

                Paragraph p7 = new Paragraph(string.Format("Este relatório contém dados e informações referentes aos atendimentos de nossas equipes, tanto de desenvolvimento quanto de suporte de aplicações, registrados em nosso sistema de suporte para o cliente Avaya Brasil  referentes ao mês de " + referencia1 + " de " + referencia2 + "."), fonte5);
                doc.Add(p7);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p7);
                doc.Add(new Paragraph(" ")); //Quebra de linha

                Chunk p8 = new Chunk(string.Format("Os chamados podem ser visualizados individualmente através do acesso WEB ao site do Suporte MidiaVox, endereço:"), fonte5);
                Chunk p81 = new Chunk(string.Format(" http://suporte.midiavox.com.br"), fonte6);
                Phrase frase1 = new Phrase();
                frase1.Add(p8);
                frase1.Add(p81);
                doc.Add(new Paragraph(frase1));
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + frase1);

                doc.Add(new Paragraph(" ")); //Quebra de linha


                string DataDadosExtraidos = DateTime.Now.ToString("dd/MM/yyyy");
                //string DataDadosExtraidos = dataNow.ToString("dd/MM/yyyy");
                
                Paragraph p9 = new Paragraph(string.Format("Os dados foram extraídos em " + DataDadosExtraidos + " e para os chamados sem data de fechamento, a quantidade de horas trabalhadas ainda podem sofrer alterações depois desta data."), fonte5);
                doc.Add(p9);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p9);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha


                iTextSharp.text.Image ImagemPdf5 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\Minilogo.png");
                Paragraph pTextoImagem2 = new Paragraph(string.Format("Relatório Atendimento Suporte MidiaVox – Cliente AVAYA"), fonte4);
                pTextoImagem2.Alignment = Element.ALIGN_JUSTIFIED;
                doc.Add(ImagemPdf5);
                LogHelper.WriteDebugLog("adicionando Imagem ao relatório: " + ImagemPdf5);
                doc.Add(pTextoImagem2);
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pTextoImagem2);

                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha

                Chunk pNumero = new Chunk(string.Format("2."), fonte7);
                Chunk pDetalhe = new Chunk(string.Format(" Detalhes dos chamados"), fonte4);
                Phrase frase2 = new Phrase();
                frase2.Add(pNumero);
                frase2.Add(pDetalhe);
                doc.Add(new Paragraph(frase2));
                LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + frase2);
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha
                doc.Add(new Paragraph(" ")); //Quebra de linha

                //inicialização da tabela  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                string Stringdeconexao = Properties.Settings.Default.cone;
                Properties.Settings.Default.Save();
                string ServerIp = Properties.Settings.Default.svip;
                Properties.Settings.Default.Save();
                string usuario = Properties.Settings.Default.usuario;
                Properties.Settings.Default.Save();
                string senha = Properties.Settings.Default.senha;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
                LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
                LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: " + usuario);
                LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);

                DataTable dt = new DataTable();
                LogHelper.WriteDebugLog("Criando DataTable: " + dt);
                OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
                LogHelper.WriteDebugLog("abrindo conexão com a base de dados");
                //Relatorio início

                DateTime datadehoje = DateTime.Today;
                OleDbCommand comando = new OleDbCommand("select p.id as Chamado , ISNULL(SUBSTRING(p.title, 4,13), '') as SR, p.client as Cliente, SUBSTRING(p.title, 18, LEN(p.title)) as Titulo , CONVERT(CHAR(10), p.start_date, 103) as Abertura , month(p.start_date) as MesAbertura , year(p.start_date) as AnoAbertura , ISNULL(CONVERT(CHAR(10), p.close_date, 103), '') as DataFechamento , p.time_spent as HorasTrabalhadas , p.uid as ResponsavelAvaya , (select fname from tblusers where sid = p.rep) as ResponsavelMidiaVox from dbo.problems as p where p.department=1 and ( p.start_date < '" + datadehoje.Year + "-" + datadehoje.Month + "-01 00:00:00' and p.close_date is null or ( p.close_date >= '" + dataMesAnterior.Year + "-" + dataMesAnterior.Month + "-01 00:00:00' and  p.close_date < '" + datadehoje.Year + "-" + datadehoje.Month + "-01 00:00:00')) and p.time_spent > 0 order by p.id ,p.start_date desc ,p.close_date desc", conexao);

                try
                {
                    conexao.Open();
                    LogHelper.WriteDebugLog("conexão aberta.");
                    comando.Connection = conexao;
                    LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);

                    OleDbDataAdapter data = new OleDbDataAdapter(comando);
                    LogHelper.WriteDebugLog("QUERRY EXECUTADA, INFORMAÇÕES RESGATADAS DA BASE DE DADOS.");
                    data.Fill(dt);
                    LogHelper.WriteDebugLog("preenchendo data table: " + dt);
                }
                catch (Exception o)
                {
                    LogHelper.WriteDebugLog("ERRO " + o);
                    LogHelper.WriteDebugLog("ERRO CONEXÃO" + o);
                }

                try
                {

                    //inicializando tabela
                    PdfPTable tabela = new PdfPTable(11);
                    tabela.WidthPercentage = 124.5f;
                    LogHelper.WriteDebugLog("criando tabela: " + tabela);

                    // cria uma célula - será usada para cada célula abaixo
                    PdfPCell celula = new PdfPCell();
                    LogHelper.WriteDebugLog("criando celula cabeçalho" + celula);

                    //montando cabeçalho
                    var titulos = new String[] { "Chamado", "#SR", "Cliente", "Título", "Abertura", "Mês Abertura", "Ano Abertura", "Data Fechamento", "Horas Trabalhadas", "Responsável Avaya", "Responsável MidiaVox" };

                    foreach (var titulo in titulos)
                    {
                        celula.Phrase = new Phrase(titulo, fonte8);
                        celula.BackgroundColor = new BaseColor(65, 105, 225);
                        celula.HorizontalAlignment = Element.ALIGN_CENTER;
                        celula.VerticalAlignment = Element.ALIGN_MIDDLE;
                        //celula.HorizontalAlignment = 1;
                        //celula.Colspan = 12;
                        tabela.AddCell(celula);
                        LogHelper.WriteDebugLog("adicionando titulos a celula cabeçalho da tabela: " + titulo);
                        LogHelper.WriteDebugLog("adicionando celula cabeçalho ao relatório" + celula);
                    }

                    PdfPCell celula2 = new PdfPCell();
                    PdfPCell celula3 = new PdfPCell();
                    string chamado;
                    string SR;
                    string cliente;
                    string titul;
                    string abertura;
                    string mes_abertura;
                    string ano_abertura;
                    string Fechamento;
                    string Horas_Trabalhadas;
                    string Responsavel_Avaya;
                    string Responsavel_Midiavox;


                    List<string> lst = new List<string>();
                    foreach (DataRow r in dt.Rows)
                    {
                        LogHelper.WriteDebugLog("Percorrendo datatable" + dt);
                        chamado = r["Chamado"].ToString();
                        SR = r["SR"].ToString();
                        cliente = r["Cliente"].ToString();
                        titul = r["Titulo"].ToString();
                        abertura = r["Abertura"].ToString();
                        mes_abertura = r["MesAbertura"].ToString();
                        ano_abertura = r["AnoAbertura"].ToString();
                        Fechamento = r["DataFechamento"].ToString();
                        Horas_Trabalhadas = r["HorasTrabalhadas"].ToString();
                        Responsavel_Avaya = r["ResponsavelAvaya"].ToString();
                        Responsavel_Midiavox = r["ResponsavelMidiavox"].ToString();

                        var campos = new String[] { chamado, SR, cliente, titul, abertura, mes_abertura, ano_abertura, Fechamento, Horas_Trabalhadas, Responsavel_Avaya, Responsavel_Midiavox };
                        foreach (var campo in campos)

                        {
                            celula2.Phrase = new Phrase(campo, fonte9);
                            LogHelper.WriteDebugLog("criando celula corpo" + celula2);
                            celula2.BackgroundColor = new BaseColor(255, 255, 225);
                            celula2.HorizontalAlignment = Element.ALIGN_CENTER;
                            celula2.VerticalAlignment = Element.ALIGN_MIDDLE;
                            //celula.Colspan = 12;
                            tabela.AddCell(celula2);
                            LogHelper.WriteDebugLog("adicionando os campos a celula corpo da tabela" + celula2);
                            LogHelper.WriteDebugLog("adicionando celula corpo ao relatório" + celula2);
                        }

                        lst.Add(r["Chamado"].ToString() + ": " + r["SR"].ToString() + ": " + r["Cliente"].ToString() + ": " + r["Titulo"].ToString() + ": " + r["Abertura"].ToString() + ": " + r["MesAbertura"].ToString() + ": " + r["AnoAbertura"].ToString() + ": " + r["DataFechamento"].ToString() + ": " + r["HorasTrabalhadas"].ToString() + ": " + r["ResponsavelAvaya"].ToString() + ": " + r["ResponsavelMidiavox"].ToString());
                        LogHelper.WriteDebugLog("criando lista com os dados do banco na aplicação: " + lst);
                    }

                    listBox1.DataSource = lst;
                    doc.Add(tabela);
                    LogHelper.WriteDebugLog("adicionando tabela: " + tabela + " ao relatório: " + doc);
                    doc.Close();
                    LogHelper.WriteDebugLog("fechando documento" + doc);

                    string EmailRecebeRelatorio = Properties.Settings.Default.EmailRecebeRelatorio;
                    Properties.Settings.Default.Save();

                    LogHelper.WriteDebugLog("Recebendo emails parametros do App.config para enviar relatório" + EmailRecebeRelatorio);

                    string DataAtual = dataNow.ToString("MM/yyyy");
                    string MesAntes = dataMesAnterior.ToString("MMMM");
                    string AnoAntes = dataMesAnterior.ToString("yyyy");

                    string subject = "Relatórios de Tickets do Suporte - " + DataAtual;
                    string iniciomsg = ("<html><head></head><body>Prezados, bom dia. <br/><br/>");
                    string mensagem = "Segue em anexo o relat&oacute;rio de tickets do suporte relativos ao m&ecirc;s de " + MesAntes + " de " + AnoAntes + "<br/><br/>Qualquer d&uacute;vida, estamos a disposi&ccedil;&atilde;o.<br/><br/>";
                    string fimmsg = " Atenciosamente, <br/>Suporte MidiaVox.</body></html>";

                    //string mensagem2 = HttpUtility.HtmlEncode(mensagem);
                    //string recebedores = andre, luiz, avaya;

                    PSendEmail(EmailRecebeRelatorio, "suporte@midiavox.com.br", subject, ServerIp, iniciomsg + mensagem + fimmsg, caminho, true, usuario, senha);
                    LogHelper.WriteDebugLog("Enviando relatório: " + doc + "para email(s): " + EmailRecebeRelatorio);

                }
                catch (Exception tbl)
                {
                    MessageBox.Show(tbl.Message);
                    LogHelper.WriteDebugLog("ERRO AO CRIAR TABELA: " + tbl);
                }

            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);
                LogHelper.WriteDebugLog("ERROR AO CRIAR RELATÓRIO" + x);
            }

            

        }

        #endregion
        #region função Envia Email com Comparativos Anuais
        public void EnviaComparativoAnual()
            {

            string Stringdeconexao = Properties.Settings.Default.cone;
            Properties.Settings.Default.Save();
            string ServerIp = Properties.Settings.Default.svip;
            Properties.Settings.Default.Save();
            string usuario = Properties.Settings.Default.usuario;
            Properties.Settings.Default.Save();
            string senha = Properties.Settings.Default.senha;
            Properties.Settings.Default.Save();
            string destinatario = Properties.Settings.Default.DestinatarioComparativo;
            Properties.Settings.Default.Save();

            LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
            LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
            LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: " + usuario);
            LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);
            LogHelper.WriteDebugLog("recebendo destinatario do App.Config: " + destinatario);

            int anoatual = DateTime.Now.Year;
            int anoanterior = anoatual - 1;
            int anoanterior2 = anoatual - 2;

            listBox1.Items.Add(anoanterior);
            listBox1.Items.Add(anoanterior2);

            DataTable dt = new DataTable();
            LogHelper.WriteDebugLog("Criando DataTable: " + dt);
            OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
            LogHelper.WriteDebugLog("abrindo conexão com a base de dados");
            //Relatorio início

            DateTime datadehoje = DateTime.Today;
            OleDbCommand comando = new OleDbCommand("select '" + anoanterior2 + "-" + anoanterior + "' as ANO, SUM(p.time_spent) as Horas, COUNT(*) as Qtd from dbo.problems as p where p.department=1 and p.start_date >= '" + anoanterior2 + "-10-01 00:00:00' and p.start_date <= '" + anoanterior + "-09-30 23:59:59' UNION select '" + anoanterior + "-" + anoatual + "' as ANO, SUM(p.time_spent) as Horas, COUNT(*) as Qtd from dbo.problems as p where p.department=1 and p.start_date >= '" + anoanterior + "-10-01 00:00:00' and p.start_date <= '" + anoatual + "-09-30 23:59:59'", conexao);

            try
            {
                conexao.Open();
                LogHelper.WriteDebugLog("conexão aberta.");
                comando.Connection = conexao;
                LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);

                OleDbDataAdapter data = new OleDbDataAdapter(comando);
                LogHelper.WriteDebugLog("QUERRY EXECUTADA, INFORMAÇÕES RESGATADAS DA BASE DE DADOS.");
                data.Fill(dt);
                LogHelper.WriteDebugLog("preenchendo data table: " + dt);
            }
            catch (Exception o)
            {
                LogHelper.WriteDebugLog("ERRO " + o);
                LogHelper.WriteDebugLog("ERRO CONEXÃO" + o);
            }

            try
            {
                string mensagem = "<table border=7><tr><th>Ano</th><th>Horas</th><th>Qtd</th></tr>";
                string Horas;
                string Ano;
                string data2 = DateTime.Now.ToString(new CultureInfo("pt-BR", false).DateTimeFormat.ShortDatePattern);
                string qtd;
                string subject = "Comparação entre " + anoanterior2 + "-" + anoanterior + " e " + anoanterior + "-" + anoatual;
                string iniciomsg = "<html><head></head><body>Ol&aacute;, segue abaixo o comparativo entre os per&iacute;odos  " + anoanterior2 + "-" + anoanterior + " e " + anoanterior + "-" + anoatual + ". <br/><br/>";
                string fimmsg = "</tbody></table><br/> Atenciosamente, <br/>Suporte MidiaVox.</body></html>";

                List<string> lst = new List<string>();
                foreach (DataRow r in dt.Rows)
                {
                    LogHelper.WriteDebugLog("Percorrendo datatable" + dt);

                    Ano = r["ANO"].ToString();
                    Horas = r["Horas"].ToString();
                    qtd = r["Qtd"].ToString();
                    mensagem += "<tr><td>" + Ano + "</td><td>" + Horas + "</td><td>" + qtd + "</td></tr>";
                    LogHelper.WriteDebugLog("obtendo acesso a os dados: " + Ano + ", " + Horas + " e " + qtd);

                    lst.Add(r["ANO"].ToString() + ": " + r["Horas"].ToString() + ": " + r["Qtd"].ToString());
                }

                listBox1.DataSource = lst;
                LogHelper.WriteDebugLog("adicionando as informações no listbox");

                conexao.Close();
                LogHelper.WriteDebugLog("conexão com a base de dados fechada");

                PSendEmail(destinatario, "suporte@midiavox.com.br", subject, ServerIp, iniciomsg + mensagem + fimmsg, "", true, usuario, senha);
                LogHelper.WriteDebugLog("Email enviado para: andre@midiavox.com.br com os dados resgatados da base de dados.");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Falha ao enviar email: " + ex.Message);
                LogHelper.WriteDebugLog("ERRO: " + ex);
            }

        }
        #endregion
        #region Envia Email com Comparativos Mensais
        public void EnviaComparativoMensal()
        {
            string Stringdeconexao = Properties.Settings.Default.cone;
            Properties.Settings.Default.Save();
            string ServerIp = Properties.Settings.Default.svip;
            Properties.Settings.Default.Save();
            string usuario = Properties.Settings.Default.usuario;
            Properties.Settings.Default.Save();
            string senha = Properties.Settings.Default.senha;
            Properties.Settings.Default.Save();
            string destinatario = Properties.Settings.Default.DestinatarioComparativo;
            Properties.Settings.Default.Save();

            LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
            LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
            LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: " + usuario);
            LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);
            LogHelper.WriteDebugLog("recebendo destinatario do App.Config: " + destinatario);


            int mesatual = DateTime.Now.Month;
            int mesanterior = mesatual - 1;


            int anoatual = DateTime.Now.Year;
            int anoanterior = anoatual - 1;

            listBox1.Items.Add(anoanterior);

            DataTable dt = new DataTable();
            LogHelper.WriteDebugLog("Criando DataTable: " + dt);
            OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
            LogHelper.WriteDebugLog("abrindo conexão com a base de dados");
            //Relatorio início

            DateTime datadehoje = DateTime.Today;
            OleDbCommand comando = new OleDbCommand("select '" + mesanterior + "/" + anoanterior + "' as PERIODO, SUM(p.time_spent) as Horas, COUNT(*) as Qtd from dbo.problems as p where p.department=1 and p.start_date >= '" + anoanterior + "-" + mesanterior + "-01 00:00:00' and p.start_date < '" + anoanterior + "-" + mesatual + "-01 00:00:00' UNION select '" + mesanterior + "/" + anoatual + "' as PERIODO, SUM(p.time_spent) as Horas, COUNT(*) as Qtd from dbo.problems as p where p.department=1 and p.start_date >= '" + anoatual + "-" + mesanterior + "-01 00:00:00' and p.start_date < '" + anoatual + "-" + mesatual + "-01 00:00:00' ORDER BY 1 ASC", conexao);

            try
            {
                conexao.Open();
                LogHelper.WriteDebugLog("conexão aberta.");
                comando.Connection = conexao;
                LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);

                OleDbDataAdapter data = new OleDbDataAdapter(comando);
                LogHelper.WriteDebugLog("QUERRY EXECUTADA, INFORMAÇÕES RESGATADAS DA BASE DE DADOS.");
                data.Fill(dt);
                LogHelper.WriteDebugLog("preenchendo data table: " + dt);
            }
            catch (Exception o)
            {
                LogHelper.WriteDebugLog("ERRO " + o);
                LogHelper.WriteDebugLog("ERRO CONEXÃO" + o);
            }

            try
            {
                string mensagem = "<table border=7><tr><th>Periodo</th><th>Horas</th><th>Qtd</th></tr>";
                string Horas;
                string Periodo;
                string data2 = DateTime.Now.ToString(new CultureInfo("pt-BR", false).DateTimeFormat.ShortDatePattern);
                string qtd;
                string subject = "Comparação entre os períodos de " + mesanterior + "/" + anoanterior + " e " + mesanterior + "/" + anoatual;
                string iniciomsg = "<html><head></head><body>Ol&aacute;, segue abaixo o comparativo entre os per&iacute;odos do m&ecirc;s  " + mesanterior + "/" + anoanterior + " e " + mesanterior + "/" + anoatual + ". <br/><br/>";
                string fimmsg = "</tbody></table><br/> Atenciosamente, <br/>Suporte MidiaVox.</body></html>";

                List<string> lst = new List<string>();
                foreach (DataRow r in dt.Rows)
                {
                    LogHelper.WriteDebugLog("Percorrendo datatable" + dt);

                    Periodo = r["PERIODO"].ToString();
                    Horas = r["Horas"].ToString();
                    qtd = r["Qtd"].ToString();
                    mensagem += "<tr><td>" + Periodo + "</td><td>" + Horas + "</td><td>" + qtd + "</td></tr>";
                    LogHelper.WriteDebugLog("obtendo acesso a os dados: " + Periodo + ", " + Horas + " e " + qtd);

                    lst.Add(r["PERIODO"].ToString() + ": " + r["Horas"].ToString() + ": " + r["Qtd"].ToString());
                }

                listBox1.DataSource = lst;
                LogHelper.WriteDebugLog("adicionando as informações no listbox");

                conexao.Close();
                LogHelper.WriteDebugLog("conexão com a base de dados fechada");

                PSendEmail(destinatario, "suporte@midiavox.com.br", subject, ServerIp, iniciomsg + mensagem + fimmsg, "", true, usuario, senha);
                LogHelper.WriteDebugLog("Email enviado para: " + destinatario + " com os dados resgatados da base de dados.");
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Falha ao enviar email: " + ex.Message);
                LogHelper.WriteDebugLog("ERRO: " + ex);
            }
        }
        #endregion

        private void timerSendMail_Tick(object sender, EventArgs e)
        {
            #region Timer EnviaEmailSemanal e EnviaEmailQtd
            try
            {
                LogHelper.WriteDebugLog("Timer iniciado.");
                string DiaSemana = new CultureInfo("pt-BR").DateTimeFormat.GetDayName(DateTime.Now.DayOfWeek);
                string DiaSemanaEnvioChamados = Properties.Settings.Default.diadasemanaEnvioChamados;
                Properties.Settings.Default.Save();
                string hora = DateTime.Now.ToString("HH:mm");
                string HorarioPessoais = Properties.Settings.Default.HorarioChamadosPessoais;
                Properties.Settings.Default.Save();
                string HorarioQtdChamados = Properties.Settings.Default.HorarioQtdChamados;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("timer recebendo PARAMETRO HORA do App.Config" + HorarioPessoais);
                LogHelper.WriteDebugLog("timer recebendo PARAMETRO HORA do App.Config" + HorarioQtdChamados);
                LogHelper.WriteDebugLog("Timer recebendo PARAMETRO DATA do App.Config:" + DiaSemanaEnvioChamados);

                LogHelper.WriteDebugLog("COMPARANDO PARAMETROS DO App.Config com o Horário e dia da semana atual, se forem iguais o timer será executado");
                if (DiaSemana == DiaSemanaEnvioChamados)
                {
                    LogHelper.WriteDebugLog("PARAMETRO EQUIVALENTE COM O DIA DA SEMANA ATUAL: " + DiaSemana);
                    if (hora == HorarioPessoais)
                    {
                        LogHelper.WriteDebugLog("PARAMETRO EQUIVALENTE COM O HORÁRIO ATUAL, EXECUTANDO EnviaEmailSemanal");
                        EnviaEmailSemanal();
                    }
                     if (hora== HorarioQtdChamados)
                    {
                        LogHelper.WriteDebugLog("PARAMETRO EQUIVALENTE COM O HORÁRIO ATUAL, EXECUTANDO EnviaEmailQtd");
                        EnviaEmailQtd();
                    }
                }
            }
            catch (Exception s)
            {
                LogHelper.WriteDebugLog("ERRO: " + s);
                MessageBox.Show(s.Message);
            }
            #endregion
            #region Timer AtualuzacaoChamado
            try
            {
                
                string DiaSemana = new CultureInfo("pt-BR").DateTimeFormat.GetDayName(DateTime.Now.DayOfWeek);

                Properties.Settings.Default.Save();
                string hora = DateTime.Now.ToString("HH:mm");
                string HoraAtualizacaoChamados = Properties.Settings.Default.HorarioAtualizacaoChamados;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("timer recebendo PARAMETRO HORA do App.Config: " + HoraAtualizacaoChamados);

                if (DiaSemana != "sábado" && DiaSemana != "domingo" && hora == HoraAtualizacaoChamados)
                {
                    LogHelper.WriteDebugLog("PARAMETROs EQUIVALENTEs COM O HORÁRIO E DATA ATUAL, EXECUTANDO AtualizacaoChamado");
                    AtualizacaoChamado();
                }                    
            }
            catch (Exception b)
            {
                LogHelper.WriteDebugLog("ERRO: " + b);
                MessageBox.Show(b.Message);
            }
            #endregion
            #region Timer EnviaRelatorioMensal
            try
            {
                string DiaDoMes = DateTime.Now.ToString("dd");
                string MesEnviaRelatorio = Properties.Settings.Default.MesEnviaRelatorio;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("timer recebendo PARAMETRO dia do MES do App.Config: " + DiaDoMes);

                string hora = DateTime.Now.ToString("HH:mm");
                string HoraEnviaRelatorio = Properties.Settings.Default.HoraEnviaRelatorio;
                Properties.Settings.Default.Save();
                LogHelper.WriteDebugLog("timer recebendo PARAMETRO HORA do App.Config: " + hora);

                if (DiaDoMes == MesEnviaRelatorio && hora == HoraEnviaRelatorio)
                {
                    LogHelper.WriteDebugLog("PARAMETROs EQUIVALENTEs COM O DIA DO MES E HORA ATUAL, EXECUTANDO EnviaRelatorioMensal");
                    DateTime data = DateTime.Now;
                    EnviaRelatorioMensal(data);
                }

            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);
                LogHelper.WriteDebugLog("ERRO NO TIMER NO ENVIO DO RELATORIO: " + x);
            }
            #endregion
            #region Timer EnviaComparativoAnual

            try
            {
                string AnoAtual = DateTime.Now.ToString("yyyy");
                string hora = DateTime.Now.ToString("HH:mm");
                string DataAtual = DateTime.Now.ToShortDateString();

                if (DataAtual == "1/10/"+AnoAtual && hora == "00:00")
                {
                    EnviaComparativoAnual();
                }

            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);
                LogHelper.WriteDebugLog("ERRO NO TIMER NO ENVIO DO RELATORIO: " + x);
            }

            #endregion
            #region Timer EnviaComparativoMensal
            try
            {
                string DataAtual = DateTime.Now.ToShortDateString();
                string AnoAtual = DateTime.Now.ToString("yyyy");
                string MesAtual = DateTime.Now.ToString("MM");
                string hora = DateTime.Now.ToString("HH:mm");

                if (DataAtual=="01"+"/"+MesAtual+"/"+AnoAtual && hora=="00:00")
                {
                    EnviaComparativoMensal();
                }
            }
            catch (Exception x)
            {

              
            }
            #endregion
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
               
                string hora = DateTime.Now.ToString("H:mm");
                listBox1.Items.Add(hora);

                
            }catch(Exception x)
            {
                MessageBox.Show(x.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region Botão Chamados atrasados
            AtualizacaoChamado();
            #endregion
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            DateTime data = DateTime.Now;
            EnviaRelatorioMensal(data);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {

        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            #region Envia relatório do mês selecionado

            string dataSelecionada = (string)comboBox1.SelectedItem;
            //listBox1.Items.Add(dataSelecionada);

            string MesSelecionadoString = dataSelecionada.Substring(0, 1 + 1);
            int MesSelecionadoInt = Convert.ToInt32(MesSelecionadoString);
            string MesSelecionadoExtenso = new DateTime(1900, MesSelecionadoInt, 1).ToString("MMMM", new CultureInfo("pt-BR"));

            DateTime MesSelecionadoDateTime = DateTime.ParseExact(MesSelecionadoString, "MM", CultureInfo.InvariantCulture);


            string AnoSelecionado = dataSelecionada.Substring(1 + 2);
            DateTime AnoSelecionadoDateTime = DateTime.ParseExact(MesSelecionadoString + "/" + AnoSelecionado + " 00:00", "MM/yyyy HH:mm", CultureInfo.InvariantCulture);

            EnviaRelatorioMensal(AnoSelecionadoDateTime);

        //    try
        //    {
        //        Document doc = new Document(PageSize.A4);//criando e estipulando o tipo da folha usada
        //        doc.SetMargins(60, 60, 60, 60);//estibulando o espaçamento das margens que queremos
        //        doc.AddCreationDate();//adicionando as configuracoes

        //        //caminho onde sera criado o pdf + nome desejado
        //        //OBS: o nome sempre deve ser terminado com .pdf
        //        //DateTime dataMesAnterior = DateTime.Now;
        //        //dataMesAnterior = dataMesAnterior.AddMonths(-1);
        //        //string MesDoRelatorio = dataMesAnterior.ToString("MMMM - yyyy");
        //        DateTime dataMesAnterior = DateTime.Now;
        //        dataMesAnterior = dataMesAnterior.AddMonths(-1);
        //        string MesDoRelatorio = dataMesAnterior.ToString("MMMM - yyyy");

        //        string dataSelecionada = (string)comboBox1.SelectedItem;
        //        //listBox1.Items.Add(dataSelecionada);

        //        string MesSelecionadoString = dataSelecionada.Substring(0, 1 + 1);
        //        int MesSelecionadoInt = Convert.ToInt32(MesSelecionadoString);                      
        //        string MesSelecionadoExtenso = new DateTime(1900, MesSelecionadoInt, 1).ToString("MMMM", new CultureInfo("pt-BR"));

        //        DateTime MesSelecionadoDateTime = DateTime.ParseExact(MesSelecionadoString, "MM", CultureInfo.InvariantCulture);


        //        string AnoSelecionado = dataSelecionada.Substring(1 + 2);
        //        DateTime AnoSelecionadoDateTime = DateTime.ParseExact(MesSelecionadoString +"/"+ AnoSelecionado + " 00:00", "MM/yyyy HH:mm", CultureInfo.InvariantCulture);
        //        listBox1.Items.Add(AnoSelecionadoDateTime);
        //        listBox1.Items.Add(AnoSelecionadoDateTime.Month);
        //        listBox1.Items.Add(AnoSelecionadoDateTime.Year);



        //        String pathApp = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
        //        string caminho = pathApp + "/*Relatorios*/\\Relatório de Atendimento Suporte MidiaVox - Avaya - " + dataSelecionada.Replace("/", " - ") + ".pdf";

        //        DateTime dataAtual = DateTime.Now;



        //        LogHelper.WriteDebugLog("Adicionando documento ao caminho: " + caminho);

        //        //criando o arquivo pdf embranco, passando como parametro a variavel                
        //        //doc criada acima e a variavel caminho 
        //        //tambem criada acima.
        //        PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(caminho, FileMode.Create));

        //        doc.Open();
        //        LogHelper.WriteDebugLog("Abrindo documento: " + doc);

        //        //Cores e Fontes
        //        BaseColor AzulMeiaNoite = new BaseColor(25, 25, 112);
        //        iTextSharp.text.Font fonte1 = FontFactory.GetFont("Calibri", 36, AzulMeiaNoite);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + AzulMeiaNoite + " " + fonte1);

        //        BaseColor RoyalBlue = new BaseColor(65, 105, 225);
        //        iTextSharp.text.Font fonte2 = FontFactory.GetFont("Calibri", 20, RoyalBlue);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + RoyalBlue + " " + fonte2);

        //        BaseColor CinzaEscuro = new BaseColor(105, 105, 105);
        //        iTextSharp.text.Font fonte3 = FontFactory.GetFont("Calibri", 16, CinzaEscuro);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + CinzaEscuro + " " + fonte3);

        //        iTextSharp.text.Font fonte4 = FontFactory.GetFont("Calibri", 11);
        //        LogHelper.WriteDebugLog("criando fonte: " + fonte4);

        //        BaseColor VerdeOliva = new BaseColor(107, 142, 35);
        //        iTextSharp.text.Font fonte5 = FontFactory.GetFont("Calibri", 14, VerdeOliva);
        //        iTextSharp.text.Font fonte51 = FontFactory.GetFont("Calibri", 28, VerdeOliva);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + VerdeOliva + " " + fonte5);
        //        LogHelper.WriteDebugLog("criando fonte: " + fonte51);


        //        BaseColor AzulLink = new BaseColor(0, 0, 255);
        //        iTextSharp.text.Font fonte6 = FontFactory.GetFont("Calibri", 14, AzulLink);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + AzulLink + " " + fonte6);

        //        BaseColor Preto = new BaseColor(0, 0, 0);
        //        iTextSharp.text.Font fonte7 = FontFactory.GetFont("Liberation Serif", 14, Preto);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + Preto + " " + fonte7);

        //        BaseColor Branco = new BaseColor(255, 255, 255);
        //        iTextSharp.text.Font fonte8 = FontFactory.GetFont("Calibri", 8, Branco);
        //        LogHelper.WriteDebugLog("criando cor e fonte: " + Branco + " " + fonte8);

        //        iTextSharp.text.Font fonte9 = FontFactory.GetFont("Calibri", 7, Preto);
        //        LogHelper.WriteDebugLog("criando fonte: " + fonte9);
        //        //

        //        string Imagem = pathApp + "Imagens\\logo MIDIAVOX.jpg";
        //        iTextSharp.text.Image ImagemPdf = iTextSharp.text.Image.GetInstance(Imagem);
        //        doc.Add(ImagemPdf);
        //        LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf);

        //        //Parágrafos
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        Paragraph p1 = new Paragraph(string.Format("Relatório de Atendimento"), fonte1);
        //        doc.Add(p1);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p1);

        //        Paragraph p2 = new Paragraph(string.Format("Cliente: AVAYA BRASIL"), fonte2);
        //        doc.Add(p2);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p2);

        //        Paragraph p3 = new Paragraph(string.Format("Gerência de Suporte - MidiaVox"), fonte3);
        //        doc.Add(p3);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p3);

        //        Paragraph p4 = new Paragraph(string.Format("Referência: " + dataSelecionada.Replace("/", "-")), fonte3);
        //        doc.Add(p4);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p4);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        iTextSharp.text.Image ImagemPdf2 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\onda.png");
        //        doc.Add(ImagemPdf2);
        //        LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf2);

        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        iTextSharp.text.Image ImagemPdf3 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\Minilogo.png");
        //        Paragraph p5 = new Paragraph(string.Format("Relatório Atendimento Suporte MidiaVox – Cliente AVAYA"), fonte4);
        //        p5.Alignment = Element.ALIGN_JUSTIFIED;
        //        doc.Add(ImagemPdf3);
        //        LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf3);
        //        doc.Add(p5);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p5);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        Paragraph pSum = new Paragraph(string.Format("                      SUMÁRIO"), fonte51);
        //        doc.Add(pSum);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pSum);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        Paragraph pSum1 = new Paragraph(string.Format("            Apresentação........................................................................03"), fonte5);
        //        doc.Add(pSum1);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pSum1);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        Paragraph pSum2 = new Paragraph(string.Format("            Detalhes dos chamados........................................................04"), fonte5);
        //        doc.Add(pSum2);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pSum2);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        iTextSharp.text.Image ImagemPdf4 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\Minilogo.png");
        //        Paragraph pTextoImagem = new Paragraph(string.Format("Relatório Atendimento Suporte MidiaVox – Cliente AVAYA"), fonte4);
        //        pTextoImagem.Alignment = Element.ALIGN_JUSTIFIED;
        //        doc.Add(ImagemPdf4);
        //        LogHelper.WriteDebugLog("adicionando imagem ao relatório: " + ImagemPdf4);
        //        doc.Add(pTextoImagem);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pTextoImagem);

        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        Paragraph p6 = new Paragraph(string.Format("1. Apresentação "), fonte5);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(p6);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p6);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

                

        //        Paragraph p7 = new Paragraph(string.Format("Este relatório contém dados e informações referentes aos atendimentos de nossas equipes, tanto de desenvolvimento quanto de suporte de aplicações, registrados em nosso sistema de suporte para o cliente Avaya Brasil  referentes ao mês de " + MesSelecionadoExtenso + " de " + AnoSelecionado + "."), fonte5);
        //        doc.Add(p7);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p7);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        Chunk p8 = new Chunk(string.Format("Os chamados podem ser visualizados individualmente através do acesso WEB ao site do Suporte MidiaVox, endereço:"), fonte5);
        //        Chunk p81 = new Chunk(string.Format(" http://suporte.midiavox.com.br"), fonte6);
        //        Phrase frase1 = new Phrase();
        //        frase1.Add(p8);
        //        frase1.Add(p81);
        //        doc.Add(new Paragraph(frase1));
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + frase1);

        //        doc.Add(new Paragraph(" ")); //Quebra de linha


        //        string DataDadosExtraidos = DateTime.Now.ToString("dd/MM/yyyy");
        //        Paragraph p9 = new Paragraph(string.Format("Os dados foram extraídos em " + DataDadosExtraidos + " e para os chamados sem data de fechamento, a quantidade de horas trabalhadas ainda podem sofrer alterações depois desta data."), fonte5);
        //        doc.Add(p9);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + p9);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha


        //        iTextSharp.text.Image ImagemPdf5 = iTextSharp.text.Image.GetInstance(pathApp + "Imagens\\Minilogo.png");
        //        Paragraph pTextoImagem2 = new Paragraph(string.Format("Relatório Atendimento Suporte MidiaVox – Cliente AVAYA"), fonte4);
        //        pTextoImagem2.Alignment = Element.ALIGN_JUSTIFIED;
        //        doc.Add(ImagemPdf5);
        //        LogHelper.WriteDebugLog("adicionando Imagem ao relatório: " + ImagemPdf5);
        //        doc.Add(pTextoImagem2);
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + pTextoImagem2);

        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        Chunk pNumero = new Chunk(string.Format("2."), fonte7);
        //        Chunk pDetalhe = new Chunk(string.Format(" Detalhes dos chamados"), fonte4);
        //        Phrase frase2 = new Phrase();
        //        frase2.Add(pNumero);
        //        frase2.Add(pDetalhe);
        //        doc.Add(new Paragraph(frase2));
        //        LogHelper.WriteDebugLog("adicionando parágrafo ao relatório: " + frase2);
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha
        //        doc.Add(new Paragraph(" ")); //Quebra de linha

        //        //inicialização da tabela  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //        string Stringdeconexao = Properties.Settings.Default.cone;
        //        Properties.Settings.Default.Save();
        //        string ServerIp = Properties.Settings.Default.svip;
        //        Properties.Settings.Default.Save();
        //        string usuario = Properties.Settings.Default.usuario;
        //        Properties.Settings.Default.Save();
        //        string senha = Properties.Settings.Default.senha;
        //        Properties.Settings.Default.Save();
        //        LogHelper.WriteDebugLog("recebendo STRING DE CONEXÃO do App.Config: " + Stringdeconexao);
        //        LogHelper.WriteDebugLog("recebendo IP DO SERVIDOR do App.Config: " + ServerIp);
        //        LogHelper.WriteDebugLog("recebendo USUARIO do App.Config: " + usuario);
        //        LogHelper.WriteDebugLog("recebendo senha do App.Config: " + senha);

        //        DataTable dt = new DataTable();
        //        LogHelper.WriteDebugLog("Criando DataTable: " + dt);
        //        OleDbConnection conexao = new OleDbConnection(Stringdeconexao);
        //        LogHelper.WriteDebugLog("abrindo conexão com a base de dados");
        //        //Relatorio início

        //        DateTime datadehoje = DateTime.Today;
        //        OleDbCommand comando = new OleDbCommand("select p.id as Chamado , ISNULL(SUBSTRING(p.title, 4,13), '') as SR, p.client as Cliente, SUBSTRING(p.title, 18, LEN(p.title)) as Titulo , CONVERT(CHAR(10), p.start_date, 103) as Abertura , month(p.start_date) as MesAbertura , year(p.start_date) as AnoAbertura , ISNULL(CONVERT(CHAR(10), p.close_date, 103), '') as DataFechamento , p.time_spent as HorasTrabalhadas , p.uid as ResponsavelAvaya , (select fname from tblusers where sid = p.rep) as ResponsavelMidiaVox from dbo.problems as p where p.department=1 and ( p.start_date < '" + datadehoje.Year + "-" + datadehoje.Month + "-01 00:00:00' and p.close_date is null or ( p.close_date >= '" + dataMesAnterior.Year + "-" + dataMesAnterior.Month + "-01 00:00:00' and  p.close_date < '" + datadehoje.Year + "-" + datadehoje.Month + "-01 00:00:00')) and p.time_spent > 0 order by p.id ,p.start_date desc ,p.close_date desc", conexao);

        //        try
        //        {
        //            conexao.Open();
        //            LogHelper.WriteDebugLog("conexão aberta.");
        //            comando.Connection = conexao;
        //            LogHelper.WriteDebugLog("EXECUTANDO QUERRY: " + comando);

        //            OleDbDataAdapter data = new OleDbDataAdapter(comando);
        //            LogHelper.WriteDebugLog("QUERRY EXECUTADA, INFORMAÇÕES RESGATADAS DA BASE DE DADOS.");
        //            data.Fill(dt);
        //            LogHelper.WriteDebugLog("preenchendo data table: " + dt);
        //        }
        //        catch (Exception o)
        //        {
        //            LogHelper.WriteDebugLog("ERRO " + o);
        //            LogHelper.WriteDebugLog("ERRO CONEXÃO" + o);
        //        }

        //        try
        //        {

        //            //inicializando tabela
        //            PdfPTable tabela = new PdfPTable(11);
        //            tabela.WidthPercentage = 124.5f;
        //            LogHelper.WriteDebugLog("criando tabela: " + tabela);

        //            // cria uma célula - será usada para cada célula abaixo
        //            PdfPCell celula = new PdfPCell();
        //            LogHelper.WriteDebugLog("criando celula cabeçalho" + celula);

        //            //montando cabeçalho
        //            var titulos = new String[] { "Chamado", "#SR", "Cliente", "Título", "Abertura", "Mês Abertura", "Ano Abertura", "Data Fechamento", "Horas Trabalhadas", "Responsável Avaya", "Responsável MidiaVox" };

        //            foreach (var titulo in titulos)
        //            {
        //                celula.Phrase = new Phrase(titulo, fonte8);
        //                celula.BackgroundColor = new BaseColor(65, 105, 225);
        //                celula.HorizontalAlignment = Element.ALIGN_CENTER;
        //                celula.VerticalAlignment = Element.ALIGN_MIDDLE;
        //                //celula.HorizontalAlignment = 1;
        //                //celula.Colspan = 12;
        //                tabela.AddCell(celula);
        //                LogHelper.WriteDebugLog("adicionando titulos a celula cabeçalho da tabela: " + titulo);
        //                LogHelper.WriteDebugLog("adicionando celula cabeçalho ao relatório" + celula);
        //            }

        //            PdfPCell celula2 = new PdfPCell();
        //            PdfPCell celula3 = new PdfPCell();
        //            string chamado;
        //            string SR;
        //            string cliente;
        //            string titul;
        //            string abertura;
        //            string mes_abertura;
        //            string ano_abertura;
        //            string Fechamento;
        //            string Horas_Trabalhadas;
        //            string Responsavel_Avaya;
        //            string Responsavel_Midiavox;


        //            List<string> lst = new List<string>();
        //            foreach (DataRow r in dt.Rows)
        //            {
        //                LogHelper.WriteDebugLog("Percorrendo datatable" + dt);
        //                chamado = r["Chamado"].ToString();
        //                SR = r["SR"].ToString();
        //                cliente = r["Cliente"].ToString();
        //                titul = r["Titulo"].ToString();
        //                abertura = r["Abertura"].ToString();
        //                mes_abertura = r["MesAbertura"].ToString();
        //                ano_abertura = r["AnoAbertura"].ToString();
        //                Fechamento = r["DataFechamento"].ToString();
        //                Horas_Trabalhadas = r["HorasTrabalhadas"].ToString();
        //                Responsavel_Avaya = r["ResponsavelAvaya"].ToString();
        //                Responsavel_Midiavox = r["ResponsavelMidiavox"].ToString();

        //                var campos = new String[] { chamado, SR, cliente, titul, abertura, mes_abertura, ano_abertura, Fechamento, Horas_Trabalhadas, Responsavel_Avaya, Responsavel_Midiavox };
        //                foreach (var campo in campos)

        //                {
        //                    celula2.Phrase = new Phrase(campo, fonte9);
        //                    LogHelper.WriteDebugLog("criando celula corpo" + celula2);
        //                    celula2.BackgroundColor = new BaseColor(255, 255, 225);
        //                    celula2.HorizontalAlignment = Element.ALIGN_CENTER;
        //                    celula2.VerticalAlignment = Element.ALIGN_MIDDLE;
        //                    //celula.Colspan = 12;
        //                    tabela.AddCell(celula2);
        //                    LogHelper.WriteDebugLog("adicionando os campos a celula corpo da tabela" + celula2);
        //                    LogHelper.WriteDebugLog("adicionando celula corpo ao relatório" + celula2);
        //                }

        //                lst.Add(r["Chamado"].ToString() + ": " + r["SR"].ToString() + ": " + r["Cliente"].ToString() + ": " + r["Titulo"].ToString() + ": " + r["Abertura"].ToString() + ": " + r["MesAbertura"].ToString() + ": " + r["AnoAbertura"].ToString() + ": " + r["DataFechamento"].ToString() + ": " + r["HorasTrabalhadas"].ToString() + ": " + r["ResponsavelAvaya"].ToString() + ": " + r["ResponsavelMidiavox"].ToString());
        //                LogHelper.WriteDebugLog("criando lista com os dados do banco na aplicação: " + lst);
        //            }

        //            listBox1.DataSource = lst;
        //            doc.Add(tabela);
        //            LogHelper.WriteDebugLog("adicionando tabela: " + tabela + " ao relatório: " + doc);
        //            doc.Close();
        //            LogHelper.WriteDebugLog("fechando documento" + doc);

        //            //string EmailRecebeRelatorio = Properties.Settings.Default.EmailRecebeRelatorio;
        //            //Properties.Settings.Default.Save();

        //            //LogHelper.WriteDebugLog("Recebendo emails parametros do App.config para enviar relatório" + EmailRecebeRelatorio);

        //            //string DataAtual = DateTime.Now.ToString("MM/yyyy");
        //            //string MesAntes = dataMesAnterior.ToString("MMMM");
        //            //string AnoAntes = dataMesAnterior.ToString("yyyy");

        //            //string subject = "Relatórios de Tickets do Suporte - " + DataAtual;
        //            //string iniciomsg = ("<html><head></head><body>Prezados, bom dia. <br/><br/>");
        //            //string mensagem = "Segue em anexo o relat&oacute;rio de tickets do suporte relativos ao m&ecirc;s de " + MesAntes + " de " + AnoAntes + "<br/><br/>Qualquer d&uacute;vida, estamos a disposi&ccedil;&atilde;o.<br/><br/>";
        //            //string fimmsg = " Atenciosamente, <br/>Suporte MidiaVox.</body></html>";

        //            ////string mensagem2 = HttpUtility.HtmlEncode(mensagem);
        //            ////string recebedores = andre, luiz, avaya;

        //            //PSendEmail(EmailRecebeRelatorio, "suporte@midiavox.com.br", subject, ServerIp, iniciomsg + mensagem + fimmsg, caminho, true, usuario, senha);
        //            //LogHelper.WriteDebugLog("Enviando relatório: " + doc + "para email(s): " + EmailRecebeRelatorio);

        //        }
        //        catch (Exception tbl)
        //        {
        //            MessageBox.Show(tbl.Message);
        //            LogHelper.WriteDebugLog("ERRO AO CRIAR TABELA: " + tbl);
        //        }

        //    }
        //    catch (Exception x)
        //    {
        //        MessageBox.Show(x.Message);
        //        LogHelper.WriteDebugLog("ERROR AO CRIAR RELATÓRIO" + x);
        //   }


            #endregion
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            EnviaComparativoAnual();
        }

        private void button5_Click(object sender, EventArgs e)

        {
            EnviaComparativoMensal();
        }
    }

    #region Classe Log4Net
    public class LogHelper
    {
        private static readonly ILog _debugLogger;
        private static ILog GetLogger(string logName)
        {
            ILog log = LogManager.GetLogger(logName);
            return log;
        }

        static LogHelper()
        {
            //logger names are mentioned in <log4net> section of config file
            _debugLogger = GetLogger("MyApplicationDebugLog");
        }

        /// <summary>
        /// This method will write log in Log_USERNAME_date{yyyyMMdd}.log file
        /// </summary>
        /// <param name="message"></param>
        public static void WriteDebugLog(string message)
        {
            _debugLogger.DebugFormat(message);
        }
    }
    #endregion
}
