using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;

namespace WindowsFormsApplication1
{
    public partial class frmMain : Form
    {

        private String[] arquivos = new String[15];
        private TextBox[] caixas = new TextBox[15];
        private Label[] labels = new Label[15];
        private Button[] eraseButtons = new Button[15];
        private String[] arquivosBotoes = new String[5];
        private Button[] commonButtons = new Button[5];
        private String lastPath;
        private String strAppPath;
        private string strFilePath;
        private bool isListSaved;

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            desabilitaItens();
            preencheVetorLabels();
            preencheVetorCaixas();
            preencheVetorBotoesApagar();
            inicializaVetorArquivos();
            preencheVetorBotoesComuns();
            inicializaTelaAbrirArquivo();
            strAppPath = Application.ExecutablePath.Remove(Application.ExecutablePath.LastIndexOf("\\") + 1);
            lastPath = strAppPath;
            strFilePath = "";
            isListSaved = true;
        }

        private void inicializaTelaAbrirArquivo()
        {
            openFD.Title = "Abrir arquivo";
            openFD.FileName = "";
            openFD.InitialDirectory = lastPath;
            openFD.Filter =
                "Apresentações de slides|*.pps;*.ppt;*.ppsx;*.pptx|" +
                "Documentos de texto|*.doc;*.docx;*.rtf;*.txt|" +
                "Planilhas do Excel|*.xls;*.xlsx|" +
                "Arquivos de áudio|*.wav;*.mp3;*.wma|" +
                "Arquivos de vídeo|*.wmv|" +
                "Arquivos de imagem|*.bmp;*.gof;*.png;*.jpeg;*.jpg|" +
                "Arquivos PDF|*.pdf|Todos arquivos|*.*";
        }

        private void inicializaVetorArquivos()
        {
            for (int i = 0; i < 15; i++)
            {
                arquivos[i] = "";
                labels[i].Text = "";
                caixas[i].Text = "";
                caixas[i].Visible = false;
            }
            strFilePath = "";
            isListSaved = true;
        }

        private void desabilitaItens()
        {
            btnCommon1.Enabled = false;
            btnCommon2.Enabled = false;
            btnCommon3.Enabled = false;
            btnCommon4.Enabled = false;
            btnCommon5.Enabled = false;
        }

        private void preencheVetorCaixas()
        {
            caixas[0] = textBox1;
            caixas[1] = textBox2;
            caixas[2] = textBox3;
            caixas[3] = textBox4;
            caixas[4] = textBox5;
            caixas[5] = textBox6;
            caixas[6] = textBox7;
            caixas[7] = textBox8;
            caixas[8] = textBox9;
            caixas[9] = textBox10;
            caixas[10] = textBox11;
            caixas[11] = textBox12;
            caixas[12] = textBox13;
            caixas[13] = textBox14;
            caixas[14] = textBox15;
        }

        private void preencheVetorLabels()
        {
            labels[0] = lblName1;
            labels[1] = lblName2;
            labels[2] = lblName3;
            labels[3] = lblName4;
            labels[4] = lblName5;
            labels[5] = lblName6;
            labels[6] = lblName7;
            labels[7] = lblName8;
            labels[8] = lblName9;
            labels[9] = lblName10;
            labels[10] = lblName11;
            labels[11] = lblName12;
            labels[12] = lblName13;
            labels[13] = lblName14;
            labels[14] = lblName15;
        }

        private void preencheVetorBotoesApagar()
        {
            eraseButtons[0] = btnErase01;
            eraseButtons[1] = btnErase02;
            eraseButtons[2] = btnErase03;
            eraseButtons[3] = btnErase04;
            eraseButtons[4] = btnErase05;
            eraseButtons[5] = btnErase06;
            eraseButtons[6] = btnErase07;
            eraseButtons[7] = btnErase08;
            eraseButtons[8] = btnErase09;
            eraseButtons[9] = btnErase10;
            eraseButtons[10] = btnErase11;
            eraseButtons[11] = btnErase12;
            eraseButtons[12] = btnErase13;
            eraseButtons[13] = btnErase14;
            eraseButtons[14] = btnErase15;
        }

        private void preencheVetorBotoesComuns()
        {
            commonButtons[0] = btnCommon1;
            commonButtons[1] = btnCommon2;
            commonButtons[2] = btnCommon3;
            commonButtons[3] = btnCommon4;
            commonButtons[4] = btnCommon5;
        }
        // Recebe o nome do arquivo
        private String abrirArquivoLista()
        {
            if (!(System.IO.Directory.Exists(
                strAppPath + "Listas\\")))
                System.IO.Directory.CreateDirectory(
                    strAppPath + "Listas\\");
            openFD.Title = "Abrir lista de slides";
            openFD.FileName = "";
            openFD.InitialDirectory = strAppPath + "Listas\\";
            openFD.Filter =
                "Lista de slides|*.lst";
            if (openFD.ShowDialog() != DialogResult.Cancel)
                return openFD.FileName;
            else
                return "";

        }

        // Recebe o nome do arquivo
        private void abrirArquivo(ref string StrFullPath, ref string strFileName)
        {
            inicializaTelaAbrirArquivo();
            if (openFD.ShowDialog() != DialogResult.Cancel)
            {
                StrFullPath = openFD.FileName;
                strFileName = openFD.SafeFileName;
                lastPath = StrFullPath.Remove(StrFullPath.LastIndexOf("\\") + 1);
            }
            else
            {
                StrFullPath = "";
                strFileName = "";
            }
        }

        private String abrirArquivo()
        {
            inicializaTelaAbrirArquivo();
            if (openFD.ShowDialog() != DialogResult.Cancel)
            {
                lastPath = openFD.FileName.Remove(openFD.FileName.LastIndexOf("\\") + 1);
                return openFD.FileName;
            }
            else
                return "";
        }

        // Altera a propriedade Texto de uma textbox
        private void alteraTexto(int indiceCaixa, String arquivo, String nome)
        {
            if (arquivo != "")
            {
                isListSaved = false;
                TextBox caixaTexto = caixas[indiceCaixa];
                Label labelNome = labels[indiceCaixa];
                caixaTexto.Text = arquivo;
                labelNome.Text = nome;
                arquivos[indiceCaixa] = arquivo;
                eraseButtons[indiceCaixa].Enabled = true;
            }
        }

        // Limpa a propriedade Texto de uma textbox
        private void apagaTexto(int indiceCaixa)
        {
            isListSaved = false;
            TextBox caixaTexto = caixas[indiceCaixa];
            Label labelNome = labels[indiceCaixa];
            caixaTexto.Text = "";
            labelNome.Text = "";
            arquivos[indiceCaixa] = "";
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            string strName = ((Button)sender).Name;
            int indButton = Convert.ToInt32(strName.Substring(11, 2)) - 1;
            string strFileName = "", strFullPath = "";
            abrirArquivo(ref strFullPath, ref strFileName);
            alteraTexto(indButton, strFullPath, strFileName);
        }

        private void btnCommonOpen_Click(object sender, EventArgs e)
        {
            String fileChosen = abrirArquivo();
            if (fileChosen != "")
            {
                string strName = ((Button)sender).Name;
                int index = Convert.ToInt32(strName.Substring(13, 1)) - 1;
                commonButtons[index].Enabled = true;
                arquivosBotoes[index] = fileChosen;
            }
        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            int indice = (cmbInicial.SelectedIndex == -1 ? 0 : cmbInicial.SelectedIndex);
            StringBuilder strArquivos = new StringBuilder();
            while (indice < 15)
            {
                if (arquivos[indice] != "")
                {
                    strArquivos.AppendLine(arquivos[indice] + "\n");
                }
                indice++;
            }
            if (strArquivos.Length > 0)
            {
                executaSlides(strArquivos.ToString().Remove(strArquivos.Length-1));
            }
            else
                MessageBox.Show("Nã há arquivos a serem exibidos para esta seleção!");
        }

        private void criaPlaylist(String listaArquivos)
        {
            try
            {
                FileStream fileStream = new FileStream(
                    strAppPath +
                    "Playlist.txt", FileMode.Create, FileAccess.Write);
                StreamWriter writer = new StreamWriter(fileStream);
                writer.Flush();

                writer.BaseStream.Seek(0, SeekOrigin.Begin);
                writer.Write(listaArquivos);

                writer.Flush();
                writer.Close();
            }
            catch (Exception em)
            {
                MessageBox.Show(em.Message);
            }
        }

        private void criaBAT(String comando)
        {
            try
            {
                FileStream fileStream = new FileStream(Application.ExecutablePath.Remove(
                    Application.ExecutablePath.LastIndexOf("\\") + 1) +
                    "executar.bat", FileMode.Create, FileAccess.Write);
                StreamWriter writer = new StreamWriter(fileStream);
                writer.Flush();

                writer.BaseStream.Seek(0, SeekOrigin.Begin);
                writer.WriteLine(@"@echo off");
                writer.Flush();
                writer.WriteLine(comando);

                writer.Flush();
                writer.Close();
            }
            catch (Exception em)
            {
                MessageBox.Show(em.Message);
            }
        }

        // Executa slides de powerpoint
        private void executaSlides(String listaArquivos)
        {
            criaPlaylist(listaArquivos);
            String powerpointPath = @"""" +
                    ConfigurationManager.AppSettings["programFilesPath"] +
                    getOfficePath();
            String comando = powerpointPath + "pptview.exe\" /s /l \"" + strAppPath + "Playlist.txt\" /f";
            criaBAT(comando);
            System.Diagnostics.Process.Start(
                strAppPath + "executar.bat");
        }

        // Executa arquivos variados
        private void executaArquivo(String arquivo)
        {
            arquivo = arquivo.ToLower();
            // Slides do Powerpoint
            if ((arquivo.EndsWith(".pps")) || (arquivo.EndsWith(".ppsx")) ||
                (arquivo.EndsWith(".ppt")) || (arquivo.EndsWith(".pptx")))
            {
                executaSlides(arquivo);
            }

            // Arquivos de áudio / vídeo
            else if ((arquivo.EndsWith(".wav")) || (arquivo.EndsWith(".wma")) ||
                    (arquivo.EndsWith(".wmv")) || (arquivo.EndsWith(".mp3")))
            {
                String mPlayerPath = "\"" +
                        ConfigurationManager.AppSettings["programFilesPath"] +
                        "\\Windows Media Player\\wmplayer.exe\" ";
                String comando = mPlayerPath + "\"" + arquivo + "\" /fullscreen";
                criaBAT(comando);
                System.Diagnostics.Process.Start(@"C:\ICE\executar.bat");
            }

            // Documentos de texto
            else if ((arquivo.EndsWith(".doc")) || (arquivo.EndsWith(".docx")) ||
                    (arquivo.EndsWith(".rtf")) || (arquivo.EndsWith(".txt")))
            {
                String officePath = getOfficePath();
                if (officePath == "")
                {
                    MessageBox.Show("O caminho do Word não foi localizado!\nVerificar versão do Office!");
                }
                else
                {
                    String wordPath = "\"" +
                        ConfigurationManager.AppSettings["programFilesPath"] +
                        officePath + "winword.exe\" ";
                    String comando = wordPath + "\"" + arquivo + "\"";
                    criaBAT(comando);
                    System.Diagnostics.Process.Start(@"C:\ICE\executar.bat");
                }
            }

            // Planilhas de excel
            else if ((arquivo.EndsWith(".xls")) || (arquivo.EndsWith(".xlsx")))
            {
                String officePath = getOfficePath();
                if (officePath == "")
                {
                    MessageBox.Show("O caminho do Excel não foi localizado!\nVerificar versão do Office!");
                }
                else
                {
                    String wordPath = "\"" +
                        ConfigurationManager.AppSettings["programFilesPath"] +
                        officePath + "excel.exe\" ";
                    String comando = wordPath + "\"" + arquivo + "\"";
                    criaBAT(comando);
                    System.Diagnostics.Process.Start(@"C:\ICE\executar.bat");
                }
            }

            // Arquivos pdf
            else if (arquivo.EndsWith(".pdf"))
            {
                String pdfPath = "\"" +
                        ConfigurationManager.AppSettings["programFilesPath"] +
                        "\\Adobe\\Reader 10.0\\Reader\\acrord32.exe\" ";
                String comando = pdfPath + "\"" + arquivo + "\"";
                criaBAT(comando);
                System.Diagnostics.Process.Start(@"C:\ICE\executar.bat");
            }

            // Imagens
            else if ((arquivo.EndsWith(".gif")) || (arquivo.EndsWith(".png")) ||
                    (arquivo.EndsWith(".jpeg")) || (arquivo.EndsWith(".jpg")) ||
                    (arquivo.EndsWith(".bmp")))
            {
                String comando = "start \"Igreja de Cristo Eldorado\" \"" + arquivo + "\"";
                criaBAT(comando);
                System.Diagnostics.Process.Start(@"C:\ICE\executar.bat");
            }

            else
            {
                MessageBox.Show("Formato de arquivo não suportado!");
            }

        }

        private String getOfficePath()
        {
            if (ConfigurationManager.AppSettings["versaoOffice"].Equals("2010"))
            {
                return "\\Microsoft Office\\Office14\\";
            }
            else if (ConfigurationManager.AppSettings["versaoOffice"].Equals("2007"))
            {
                return "\\Microsoft Office\\Office12\\";
            }
            else if (ConfigurationManager.AppSettings["versaoOffice"].Equals("2003"))
            {
                return "\\Microsoft Office\\Office11\\";
            }
            else
            {
                return "";
            }
        }

        private void btnCommon_Click(object sender, EventArgs e)
        {
            string strName = ((Button)sender).Name;
            int index = Convert.ToInt32(strName.Substring(9, 1)) - 1;
            executaArquivo(arquivosBotoes[index]);
        }

        private void salvarListaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveSlideList(strFilePath == "");
        }

        private void saveSlideList(bool bGenerateNewFile)
        {
            try
            {
                if (bGenerateNewFile)
                {
                    if (!(System.IO.Directory.Exists(
                strAppPath + "Listas\\")))
                        System.IO.Directory.CreateDirectory(
                            strAppPath + "Listas\\");
                    String strHoje = DateTime.Now.ToString("yyyy_MM_dd");
                    saveFD.Title = "Salvar lista de slides";
                    saveFD.InitialDirectory = Application.ExecutablePath.Remove(
                        Application.ExecutablePath.LastIndexOf("\\") + 1) + "Listas\\";
                    saveFD.FileName = strHoje + ".lst";
                    saveFD.Filter =
                    "Listas de slides|*.lst";

                    if (saveFD.ShowDialog() != DialogResult.Cancel)
                    {
                        saveFile(saveFD.FileName);
                    }
                }
                else
                {
                    saveFile(strFilePath);
                }
            }
            catch (Exception em)
            {
                MessageBox.Show(em.Message);
            }
        }

        private void saveFile(string strFileName)
        {
            FileStream fileStream = new FileStream(@"" + strFileName, FileMode.Create, FileAccess.Write);
            StreamWriter writer = new StreamWriter(fileStream);
            writer.Flush();

            writer.BaseStream.Seek(0, SeekOrigin.Begin);
            String firstChar;
            for (int i = 0; i < 15; i++)
            {
                if (caixas[i].Text != "")
                {
                    if (i < 10)
                        firstChar = "0";
                    else
                        firstChar = "";

                    writer.WriteLine(firstChar + i.ToString() + caixas[i].Text);
                }
            }

            writer.Flush();
            writer.Close();

            strFilePath = saveFD.FileName;
            isListSaved = true;
        }

        private void abrirListaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                String fileChosen = abrirArquivoLista();
                if (fileChosen != "")
                {
                    inicializaVetorArquivos();
                    FileStream fileStream = new FileStream(@"" + fileChosen, FileMode.Open, FileAccess.Read);
                    StreamReader reader = new StreamReader(fileStream);

                    for (int i = 0; i < 15; i++)
                    {
                        caixas[i].Text = "";
                        arquivos[i] = "";
                    }

                    while (!reader.EndOfStream)
                    {
                        String linha = reader.ReadLine();
                        int indice = Convert.ToInt32(linha.Substring(0, 2));
                        caixas[indice].Text = linha.Substring(2);
                        arquivos[indice] = caixas[indice].Text;
                        labels[indice].Text = arquivos[indice].Remove(0, arquivos[indice].LastIndexOf("\\") + 1);
                    }

                    reader.Close();

                    strFilePath = fileChosen;
                    isListSaved = true;
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Não há uma lista de slides previamente salva!");
            }
        }

        private void sobreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Programa desenvolvido por Giulliano Leandro Moreira\n" +
                "para a Igreja de Cristo Eldorado.\n" +
                "O uso deste programa em um local senão para o qual ele\n" +
                "foi desenvolvido constitui crime de pirataria de software."
                );
        }

        private void frmMain_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int x = e.Location.X + this.Location.X;
                int y = e.Location.Y + this.Location.Y;
                menuAtalho.Show(x, y);
            }
        }

        private void limparListaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            inicializaVetorArquivos();
        }

        private void mni_SalvarComo_Click(object sender, EventArgs e)
        {
            saveSlideList(true);
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!isListSaved)
            {
                if (MessageBox.Show("Deseja salvar as alterações na lista de slides?", "", MessageBoxButtons.YesNo)
                    == System.Windows.Forms.DialogResult.Yes)
                {
                    // salva lista
                    saveSlideList(strFilePath == "");
                }
            }
        }

        private void btnErase_Click(object sender, EventArgs e)
        {
            string strName = ((Button)sender).Name;
            int index = Convert.ToInt32(strName.Substring(8, 2))-1;
            apagaTexto(index);
            eraseButtons[index].Enabled = false;
        }

    }

}
