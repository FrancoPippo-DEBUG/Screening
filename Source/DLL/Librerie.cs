using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using Serilog;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Cryptography.Pkcs;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;

namespace DLL
{
    public static class Costanti
    {
        public static readonly float epsilon = 0.0001f;
    }
    public static class Utils
    {
        /// <summary>
        /// Funzione che restituisce il percorso di partenza del programma
        /// </summary>
        /// <returns>Ritorna il percorso senza '\' finale</returns>
        public static string CurDir()
        {
            string currPath = Environment.CurrentDirectory.ToUpper();
            return currPath;
        }
        public static string ExtractBetween(string stringa, string compresoTra_INIZIO, string compresoTra_FINE)
        {
            if (stringa != "")
            {
                int pos = stringa.IndexOf(compresoTra_INIZIO);
                if (compresoTra_INIZIO.Contains("\r\n"))
                {
                    compresoTra_INIZIO = compresoTra_INIZIO.Replace("\r\n", "");
                }
                int lung = 0;
                lung = compresoTra_INIZIO.Length;

                int pos_1 = (pos >= 0 ? pos : 0) + lung; // pos iniziale
                if (pos_1 < 0)
                {
                    return "";
                }
                int pos_2 = 0;
                if (compresoTra_FINE == "")
                {
                    pos_2 = stringa.Length;
                }
                else
                {
                    pos_2 = stringa.IndexOf(compresoTra_FINE, pos_1 + 1); // fin dove
                    if (pos_2 < 0)
                    {
                        pos_2 = stringa.Length;
                    }
                }
                int len = pos_2 - pos_1;

                stringa = stringa.Substring(pos_1, len).Trim();
            }
            return stringa;
        }
        public static string sina(string stringa, string blocco)
        {
            if (stringa != "")
            {
                int pos_blocco = stringa.IndexOf(blocco);
                if (pos_blocco == -1)
                {
                    return (stringa);
                }
                else
                {
                    string tmp = stringa.Substring(0, pos_blocco);
                    return tmp;
                }
            }
            return (stringa);
        }
        public static string desa(string stringa, string blocco)
        {
            if (stringa != "")
            {
                int pos_blocco = stringa.IndexOf(blocco);

                if (pos_blocco == -1)
                {
                    return (stringa);
                }
                else
                {
                    string tmp = stringa.Substring(pos_blocco + blocco.Length).Trim();
                    return tmp;
                }
            }
            return (stringa);

        }
        /// <summary>
        /// Compara due stringhe eliminando i caratteri speciali
        /// </summary>
        /// <param name="stringa_1"></param>
        /// <param name="stringa_2"></param>
        /// <returns></returns>
        public static string depura(string stringa)
        {
            string ret = string.Empty;

            List<int> out_list = new List<int>();
            // CREO UN ARRAY CON I CODICI ASCII DELLA STRINGA
            byte[] lista_ascii = Encoding.ASCII.GetBytes(stringa.ToUpper());
            foreach (byte ascii in lista_ascii)
            {
                // FILTRO SOLO LE LETTERE MAIUSCOLE
                if (ascii > 64 & ascii < 91)
                {
                    out_list.Add(ascii);
                }
            }
            int indice = 0;
            byte[] out_byte = new byte[out_list.Count];
            foreach (int i in out_list)
            {
                out_byte[indice] = (byte)i;
                indice += 1;
            }
            // RICONVERTO I CODICI ASCII IN STRINGA
            ret = Encoding.ASCII.GetString(out_byte);

            return ret;
        }
        public static bool check_cf(string cf)
        {
            if (cf.Length != 11 | cf.StartsWith("0000000"))
            {
                return false;
            }

            int tmp = 0;
            int totn = 0;
            for (int i = 0; i < 10; i++)
            {
                tmp = int.Parse(cf.Substring(i, 1));

                if ((i + 1) % 2 != 0)
                {
                    totn += tmp;
                }
                else
                {
                    if (tmp * 2 > 9)
                    {
                        totn += (tmp * 2) - 9;
                    }
                    else
                    {
                        totn += (tmp * 2);
                    }
                }
            }
            tmp = 10 - int.Parse(totn.ToString().Substring(totn.ToString().Length - 1));

            if (tmp.ToString().Substring(tmp.ToString().Length - 1) == cf.Substring(cf.Length - 1))
            {
                return true;
            }
            return false;
        }
        public static bool comparacat(string stringa_1, string stringa_2)
        {
            int indice_stringa = 0;
            int len_stringa = 0;
            int differenze = 0;

            if (stringa_1.Length > stringa_2.Length)
            {
                len_stringa = stringa_2.Length;
            }
            else
            {
                len_stringa = stringa_1.Length;
            }

            while (indice_stringa < len_stringa)
            {
                if (!(stringa_1.Substring(indice_stringa, 1) == stringa_2.Substring(indice_stringa, 1)))
                {
                    differenze += 1;
                }
                indice_stringa += 1;
            }

            if (differenze >= 2)
            {
                return false;
            }

            return true;
        }
        public static int quante_parole(string stringa)
        {
            List<string> parole = new List<string>();
            stringa += " ";

            while (true)
            {
                if (stringa.Contains(" "))
                {
                    parole.Add(stringa.Substring(0, stringa.IndexOf(" ")));
                    stringa = stringa.Remove(0, stringa.IndexOf(" ") + 1);
                }
                else
                {
                    break;
                }
            }
            return parole.Count;
        }
        public static List<string> dividi_parole(string stringa)
        {
            List<string> result = new List<string>();
            stringa += " ";

            while (true)
            {
                if (stringa.Contains(" "))
                {
                    result.Add(stringa.Substring(0, stringa.IndexOf(" ")));
                    stringa = stringa.Remove(0, stringa.IndexOf(" ") + 1);
                }
                else
                {
                    break;
                }
            }
            return result;
        }
        public static string RimuoviNDG(string denominazione)
        {
            string ret = denominazione;
            string[] natura_giuridica = new string[] { "SRLS", "S.R.L.S.", "SRL", "S.R.L.", " SAS", "S.A.S.", "SNC", "S.N.C", " SPA", "S.P.A." };

            foreach (string ndg in natura_giuridica)
            {
                ret = ret.Replace(ndg, "");
            }

            return ret.Trim();
        }
        public static bool risolviCF(string cf, string denominazione, out string cognome, out string nome)
        {
            if (denominazione.Contains("\'"))
            {
                denominazione = denominazione.Replace("\'", "");
            }
            List<string> parole = dividi_parole(denominazione);
            cognome = string.Empty;
            nome = string.Empty;
            //data_nascita = string.Empty;

            int tentativi = 0;
            string tmp = string.Empty;
            while (true)
            {
                tentativi += 1;
                foreach (string parola in parole)
                {
                    tmp += parola;
                    if (parola.Length <= 3 & parole.Count > 2)
                    {
                        tmp += "_";
                        continue;
                    }

                    List<string> consonanti = new List<string>();
                    List<string> vocali = new List<string>();
                    foreach (char c in tmp.ToUpper())
                    {
                        if (c == 'A' | c == 'E' | c == 'I' | c == 'O' | c == 'U')
                        {
                            vocali.Add(c.ToString());
                        }
                        else if (c != '_')
                        {
                            consonanti.Add(c.ToString());
                        }
                    }

                    string ret = string.Empty;
                    int count = 0;

                    if (cognome == String.Empty)
                    {
                        // CONTROLLO COGNOME
                        switch (consonanti.Count)
                        {
                            case int quanti when quanti >= 3: // PRENDO LE PRIME TRE CONSONANTI
                                ret += consonanti[0] + consonanti[1] + consonanti[2];
                                break;

                            case int quanti when quanti < 3: // PRENDO LE PRIME CONSONANTI E USO LE RIMANTI VOCALI
                                for (int i = 0; count < 3; i++)
                                {
                                    if (i < consonanti.Count)
                                    {
                                        ret += consonanti[i];
                                        count += 1;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                for (int i = 0; count < 3; i++)
                                {
                                    if (i < vocali.Count)
                                    {
                                        ret += vocali[i];
                                        count += 1;
                                    }
                                    else
                                    {
                                        ret += "X";
                                        count += 1;
                                    }
                                }
                                break;
                        }
                        if (ret == cf.Substring(0, 3))
                        {
                            cognome = tmp.Replace("_", " ");
                            foreach (string s in dividi_parole(cognome))
                            {
                                parole.Remove(s);
                            }
                            tmp = string.Empty;
                            break;
                        }
                        else
                        {
                            ret = string.Empty;
                        }
                    }

                    if (nome == String.Empty)
                    {
                        // CONTROLLO NOME
                        switch (consonanti.Count)
                        {
                            case int quanti when quanti >= 4: // PRENDO LA PRIMA LA TERZA E LA QUARTA CONSONANTE
                                ret += consonanti[0] + consonanti[2] + consonanti[3];
                                break;

                            case int quanti when quanti == 3: // PRENDO LE PRIME TRE CONSONANTI
                                ret += consonanti[0] + consonanti[1] + consonanti[2];
                                break;

                            case int quanti when quanti < 3: // PRENDO LE PRIME CONSONANTI E USO LE RIMANTI VOCALI
                                for (int i = 0; count < 3; i++)
                                {
                                    if (i < consonanti.Count)
                                    {
                                        ret += consonanti[i];
                                        count += 1;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                for (int i = 0; count < 3; i++)
                                {
                                    if (i < vocali.Count)
                                    {
                                        ret += vocali[i];
                                        count += 1;
                                    }
                                    else
                                    {
                                        ret += "X";
                                        count += 1;
                                    }
                                }
                                break;
                        }
                        if (ret == cf.Substring(3, 3))
                        {
                            nome = tmp.Replace("_", " ");
                            foreach (string s in dividi_parole(nome))
                            {
                                parole.Remove(s);
                            }
                            tmp = string.Empty;
                            break;
                        }
                        else
                        {
                            tmp += "_";
                        }
                    }
                }

                if (cognome != String.Empty & nome != String.Empty)
                {
                    break;
                }
                else
                {
                    if (cognome == String.Empty & nome == String.Empty | tentativi > 10)
                    {
                        return false;
                    }
                    else
                    {
                        if (tmp.Length > 0)
                        {
                            nome = tmp.Substring(0, tmp.IndexOf("_"));
                            foreach (string s in dividi_parole(nome))
                            {
                                parole.Remove(s);
                            }
                            nome = string.Empty;
                        }
                        tmp = string.Empty;

                        continue;
                    }
                }
            }
            return true;
        }
        public static void risolviCF(string cf, string denominazione, out string cognome, out string nome, out string data_nascita, out string sesso)
        {
            data_nascita = string.Empty;
            sesso = "Uomo";

            if (risolviCF(cf, denominazione, out cognome, out nome))
            {
                string anno = cf.Substring(6, 2);
                string lettera_mese = cf.Substring(8, 1);
                string giorno = cf.Substring(9, 2);

                if (int.Parse(giorno) > 40)
                {
                    giorno = (int.Parse(giorno) - 40).ToString();
                    sesso = "Donna";
                }

                string mese = string.Empty;
                switch (lettera_mese)
                {
                    case "A":
                        mese = "gennaio";
                        break;
                    case "B":
                        mese = "febbraio";
                        break;
                    case "C":
                        mese = "marzo";
                        break;
                    case "D":
                        mese = "aprile";
                        break;
                    case "E":
                        mese = "maggio";
                        break;
                    case "H":
                        mese = "giugno";
                        break;
                    case "L":
                        mese = "luglio";
                        break;
                    case "M":
                        mese = "agosto";
                        break;
                    case "P":
                        mese = "settembre";
                        break;
                    case "R":
                        mese = "ottobre";
                        break;
                    case "S":
                        mese = "novembre";
                        break;
                    case "T":
                        mese = "dicembre";
                        break;
                }

                DateTime data = DateTime.Parse(giorno + "/" + mese + "/" + anno);
                data_nascita = data.ToString("dd/MM/yyyy");
            }
        }
        public static string extractP7M(string filepath)
        {
            string nome_file = ExtractBetween(filepath.ToUpper(), "", ".P7M");
            while (true)
            {
                if (nome_file.Contains("\\"))
                {
                    nome_file = desa(nome_file, "\\");
                    continue;
                }
                else
                {
                    break;
                }
            }
            string path_file = filepath.ToUpper().Replace(nome_file.ToUpper() + ".P7M", "");
            byte[] p7m_bytes = new byte[0];
            try
            {
                p7m_bytes = Convert.FromBase64String(File.ReadAllText(filepath));
            }
            catch (Exception ex)
            {
                p7m_bytes = File.ReadAllBytes(filepath);
            }

            SignedCms cms = new SignedCms();
            cms.Decode(p7m_bytes);

            byte[] file = cms.ContentInfo.Content;
            File.WriteAllBytes(path_file + nome_file, file);
            return path_file + nome_file;
        }
        public static string File2Base64(string filepath)
        {
            string ret = "";
            ret = Convert.ToBase64String(File.ReadAllBytes(filepath));
            return ret;
        }
        public static void FileFromBase64(string base64String, string out_path)
        {
            Byte[] bytes = Convert.FromBase64String(base64String);
            File.WriteAllBytes(out_path, bytes);
        }
        public static string getIP()
        {
            foreach (NetworkInterface item in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (item.OperationalStatus == OperationalStatus.Up)
                {
                    foreach (UnicastIPAddressInformation ip in item.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            return ip.Address.ToString();
                        }
                    }
                }
            }
            throw new Exception("Non sono riuscito a stabilire un indriizzpo IP");
        }
        public static int asc(char str)
        {
            int ret = Encoding.ASCII.GetBytes(str.ToString())[0];
            return ret;
        }
        public static byte[] asc(string str)
        {
            byte[] ret = Encoding.ASCII.GetBytes(str);
            return ret;
        }
        public static tipo_elemento Clona<tipo_elemento>(tipo_elemento elemento_da_copiare)
        {
            string save_element = XamlWriter.Save(elemento_da_copiare);

            StringReader stringReader = new StringReader(save_element);
            XmlReader xmlReader = XmlReader.Create(stringReader);


            tipo_elemento elemento_clone = (tipo_elemento)XamlReader.Load(xmlReader);
            return elemento_clone;
        }

        public static string TogliAccenti(string stringa)
        {
            string ret = "";

            foreach (char c in stringa)
            {
                ret += c switch
                {
                    'À' or 'Á' => "A",
                    'È' or 'É' => "B",
                    'Ì' => "I",
                    'Ò' => "O",
                    'Ù' => "U",
                    _ => c.ToString()
                };
            }

            return ret;
        }
        public static string TogliApostrofi(string stringa)
        {
            string ret = stringa;
            int idx_save = 0;

            while (true)
            {
                int idx_pos = stringa.IndexOf("\' ", idx_save);
                if (idx_pos > 0)
                {
                    char c = (stringa.ToUpper())[idx_pos - 1];
                    if (c != 'A' && c != 'E' && c != 'I' && c != 'O' && c != 'U')
                    {
                        ret = ret.Remove(idx_pos + 1, 1);                        
                    }
                    idx_save = idx_pos + 1;
                }
                else
                {
                    break;
                }
            }

            return ret;
        }
    }
    /*
    public class Excel
    {
        private excel.Application excelApp = new excel.Application();
        private excel.Workbook workbook = null;
        private excel.Worksheet foglio;
        public Excel()
        {
            this.excelApp.Visible = false;
        }
        public bool Create(string filename, bool overwrite = false)
        {
            try
            {
                this.workbook = this.excelApp.Workbooks.Add();

                if (overwrite) excelApp.DisplayAlerts = false;
                this.workbook.SaveAs
                (
                    filename,
                    AccessMode: excel.XlSaveAsAccessMode.xlExclusive,
                    ConflictResolution: excel.XlSaveConflictResolution.xlLocalSessionChanges
                );

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public bool Open(string filename)
        {
            try
            {
                this.workbook = this.excelApp.Workbooks.Open(filename);
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }
        public bool Close()
        {
            try
            {
                this.workbook.Close(true);
                this.excelApp.Quit();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public List<List<string>> ReadAllExcel(int indice_foglio = 1)
        {
            int tentativi = 0;
            int riga = 2;
            List<List<string>> ret = [];
            MessageFrame msg = new MessageFrame(4);

            try
            {
                this.foglio = (excel.Worksheet)this.workbook.Worksheets[indice_foglio];

                while (true)
                {
                    excel.Range range = foglio.UsedRange;
                    int colon = range.Columns.Count;

                    if (foglio.Cells[riga, 1].Value.ToString() == "") break;

                    msg.scrivi("Leggo riga " + riga);
                    List<string> tmp = [];

                    for (int c = 0; c < colon; c++)
                    {
                        excel.Range cella = (excel.Range)foglio.Cells[riga, c + 1];
                        if (cella.Value != null)
                        {
                            tmp.Add(cella.Value.ToString());
                        }
                        else
                        {
                            tmp.Add("");
                        }
                    }
                    riga += 1;
                    ret.Add(tmp);
                }
            }
            catch (Exception ex)
            {
                Thread.Sleep(1000);
                tentativi += 1;
                if (tentativi > 5)
                {
                    throw new Exception("Errore nel leggere su file Excel", ex);
                }
            }

            msg.chiudi();
            return ret;
        }
        public string[] ReadExcel(int riga, int indice_foglio = 1)
        {
            int tentativi = 0;
            string[] ret = { };
            MessageFrame msg = new MessageFrame(4);

            while (true)
            {
                try
                {
                    this.foglio = (excel.Worksheet)this.workbook.Worksheets[indice_foglio];

                    excel.Range range = foglio.UsedRange;
                    int colon = range.Columns.Count;
                    if (riga >= 1)
                    {
                        msg.scrivi("Leggo riga " + riga);
                        ret = new string[colon];
                        for (int c = 0; c < colon; c++)
                        {
                            excel.Range cella = (excel.Range)foglio.Cells[riga, c + 1];
                            if (cella.Value != null)
                            {
                                ret[c] = cella.Value.ToString();
                            }
                            else
                            {
                                ret[c] = "";
                            }
                        }
                    }
                    msg.chiudi();
                    return ret;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    tentativi += 1;
                    if (tentativi > 5)
                    {
                        throw new Exception("Errore nel leggere su file Excel", ex);
                    }
                }
            }
        }
        /// <summary>
        /// Scrive una riga sul file excel aperto
        /// </summary>
        /// <param name="cosa_scrivere">Riga che andrà a scrivere</param>
        /// <param name="riga">Indice riga in base 1</param>
        /// <param name="ini_col">Indice colonna in base 1</param>
        /// <param name="indice_foglio">Indice foglio in base 1</param>
        /// <exception cref="Exception"></exception>
        public void WriteExcel(List<string> cosa_scrivere, int riga, int ini_col = 1, int indice_foglio = 1)
        {
            int tentativi = 0;
            while (true)
            {
                try
                {
                    this.foglio = (excel.Worksheet)this.workbook.Worksheets[indice_foglio];
                    excel.Range range = this.foglio.UsedRange;
                    MessageFrame msg = new MessageFrame(4);
                    msg.scrivi("Sto scrivendo il file excel");
                    int colon = cosa_scrivere.Count;
                    int i = 0;
                    if (riga >= 1)
                    {
                        for (int c = ini_col; i < colon; c++)
                        {
                            foglio.Cells[riga, c] = cosa_scrivere[i];
                            Thread.Sleep(100);
                            i += 1;
                        }
                    }
                    msg.chiudi();
                    this.workbook.Save();
                    break;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    string nome_file = this.workbook.Path + "\\" + this.workbook.Name;
                    Close();
                    excelApp.Quit();

                    excelApp = new excel.Application();
                    this.workbook = this.excelApp.Workbooks.Open(nome_file);
                    continue;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    tentativi += 1;
                    if (tentativi > 5)
                    {
                        throw new Exception("Errore nel scrivere su file Excel", ex);
                    }
                }
            }
        }
        /// <summary>
        /// Scrive un range di righe sul file excel aperto
        /// </summary>
        /// <param name="cosa_scrivere">Range di righe che andrà a scrivere</param>
        /// <param name="riga">Indice riga in base 1</param>
        /// <param name="ini_col">Indice colonna in base 1</param>
        /// <param name="indice_foglio">Indice foglio in base 1</param>
        /// <exception cref="Exception"></exception>
        public void WriteExcel(List<List<string>> cosa_scrivere, int riga, int ini_col = 1, int indice_foglio = 1)
        {
            int tentativi = 0;
            while (true)
            {
                try
                {
                    this.foglio = (excel.Worksheet)this.workbook.Worksheets[indice_foglio];
                    excel.Range range = this.foglio.UsedRange;
                    MessageFrame msg = new MessageFrame(4);
                    int righe = cosa_scrivere.Count;
                    int colon = cosa_scrivere[0].Count;

                    int i = 0;
                    if (riga >= 1)
                    {
                        for (int r = riga; r <= righe; r++)
                        {
                            msg.scrivi("");
                            msg.scrivi("Scrivo riga " + r);
                            for (int c = ini_col; i < colon; c++)
                            {
                                foglio.Cells[r, c] = cosa_scrivere[r - 1][i];
                                i += 1;
                            }
                            i = 0;
                        }
                    }
                    msg.chiudi();
                    this.workbook.Save();
                    break;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    string nome_file = this.workbook.Path + "\\" + this.workbook.Name;
                    Close();
                    excelApp.Quit();

                    excelApp = new excel.Application();
                    this.workbook = this.excelApp.Workbooks.Open(nome_file);
                    continue;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    tentativi += 1;
                    if (tentativi > 5)
                    {
                        throw new Exception("Errore nel scrivere su file Excel", ex);
                    }
                }
            }
        }
        public void NomeFoglio(int indice_foglio, string nuovo_nome)
        {
            excel.Sheets x = this.workbook.Worksheets;
            if (x.Count < indice_foglio)
            {
                this.workbook.Worksheets.Add(After: this.workbook.Sheets[x.Count]);
                //this.workbook.Worksheets[1].Move(After: this.workbook.Sheets[x.Count]);
            }
            this.foglio = (excel.Worksheet)this.workbook.Worksheets[indice_foglio];
            this.foglio.Name = nuovo_nome;
            this.workbook.Save();
        }
    }
    */
    public class Excel
    {
        private IXLWorkbook workbook = null;
        private IXLWorksheet foglio = null;
        public Excel() {}
        public bool Create(string filename, int quanti_fogli = 1, bool overwrite = false)
        {
            try
            {
                if (File.Exists(filename))
                {
                    if (!overwrite)
                    {
                        MessageBoxResult ret = MsgBox.Show("Il file " + filename + " esiste già. Sostituire?", "Sovrascrivere?", MessageBoxImage.Warning, MessageBoxButton.YesNo);
                        if (ret == MessageBoxResult.No)
                        {
                            throw new Exception("Il file " + filename + " esiste già. Impossibile continuare");
                        }
                    }
                    File.Delete(filename);
                }

                this.workbook = new XLWorkbook();
                for (int indice_foglio = 1; indice_foglio <= quanti_fogli; indice_foglio++)
                {
                    workbook.AddWorksheet("Foglio" + indice_foglio);
                }
                workbook.SaveAs(filename);

                return true;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("esiste già"))
                {
                    throw new Exception("Il file " + filename + " esiste già. Impossibile continuare");
                }
                return false;
            }
        }
        public bool Open(string filename)
        {
            try
            {
                this.workbook = new XLWorkbook(filename);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public bool Close()
        {
            try
            {
                if (this.workbook != null)
                {
                    this.workbook.Dispose();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public List<List<string>> ReadAllExcel(int indice_foglio = 1)
        {
            List<List<string>> ret = new List<List<string>>();
            int tentativi = 0;
            MessageFrame msg = new MessageFrame(4);
            try
            {
                this.foglio = this.workbook.Worksheet(indice_foglio);
                msg.scrivi("Leggo il file excel");

                IXLRange? range = foglio.RangeUsed();
                int colon = range is null ? 0 : range.ColumnsUsed().Count();
                foreach (IXLRow row in this.foglio.RowsUsed())
                {
                    List<string> tmp = [];
                    for (int c = 0; c < colon; c++)
                    {
                        IXLCell cella = row.Cell(c + 1);
                        if (cella.Value.ToString() != "")
                        {
                            tmp.Add(cella.Value.ToString());
                        }
                        else
                        {
                            tmp.Add("");
                        }
                    }
                    ret.Add(tmp);
                }
            }
            catch (Exception ex)
            {
                Thread.Sleep(1000);
                tentativi += 1;
                if (tentativi > 5)
                {
                    throw new Exception("Errore nel leggere su file Excel", ex);
                }
            }
            msg.chiudi();
            return ret;
        }
        public string[] ReadExcel(int riga, int indice_foglio = 1)
        {
            int tentativi = 0;
            string[] ret = { };
            MessageFrame msg = new MessageFrame(4);
            while (true)
            {
                try
                {
                    this.foglio = this.workbook.Worksheet(indice_foglio);
                    msg.scrivi("Leggo riga " + riga);

                    IXLRange? range = foglio.RangeUsed();
                    int colon = range is null ? 0 : range.ColumnsUsed().Count();

                    IXLRow row = this.foglio.Row(riga);

                    ret = new string[colon];
                    for (int i = 0; i < colon; i++)
                    {
                        ret[i] = row.Cell(i + 1).Value.ToString();
                    }

                    msg.chiudi();
                    return ret;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    tentativi += 1;
                    if (tentativi > 5)
                    {
                        throw new Exception("Errore nel leggere su file Excel", ex);
                    }

                }
            }
        }
        public void WriteExcel(List<string> cosa_scrivere, int riga, int ini_col = 1, int indice_foglio = 1)
        {
            int tentativi = 0;
            while (true)
            {
                try
                {
                    this.foglio = this.workbook.Worksheet(indice_foglio);
                    MessageFrame msg = new MessageFrame(4);
                    msg.scrivi("Sto scrivendo il file excel");

                    IXLRow row = this.foglio.Row(riga);
                    for (int i = 0; i < cosa_scrivere.Count; i++)
                    {
                        row.Cell(i + 1).Value = cosa_scrivere[i];
                    }

                    msg.chiudi();
                    this.workbook.Save();
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    tentativi += 1;
                    if (tentativi > 5)
                    {
                        throw new Exception("Errore nel scrivere su file Excel", ex);
                    }
                }
            }
        }
        public void WriteExcel(List<List<string>> cosa_scrivere, int riga, int ini_col = 1, int indice_foglio = 1)
        {
            int tentativi = 0;
            while (true)
            {
                try
                {
                    this.foglio = this.workbook.Worksheet(indice_foglio);
                    MessageFrame msg = new MessageFrame(4);
                    msg.scrivi("Sto scrivendo il file excel");
                    foreach (List<string> row in cosa_scrivere)
                    {
                        IXLRow excelRow = this.foglio.Row(riga);

                        int c = ini_col;
                        foreach (string cell in row)
                        {
                            excelRow.Cell(c).Value = cell;
                            c++;
                        }
                        riga++;
                    }
                    msg.chiudi();
                    this.workbook.Save();
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    tentativi += 1;
                    if (tentativi > 5)
                    {
                        throw new Exception("Errore nel scrivere su file Excel", ex);
                    }
                }
            }
        }
        public void NomeFoglio(int indice_foglio, string nuovo_nome)
        {
            if (this.workbook.Worksheets.Count < indice_foglio)
            {
                this.workbook.AddWorksheet("Foglio" + indice_foglio);
            }
            this.foglio = this.workbook.Worksheet(indice_foglio);
            this.foglio.Name = nuovo_nome;
            this.workbook.Save();
        }
}
    public static class HTML
    {
        public static ChromeDriver chrome;
        public static FirefoxDriver firefox;
        public static EdgeDriver edge;

        public static bool automate = true;
        public static string browser;

        //private LogFile log;
        public static dynamic ApriBrowser(string browser_, string link = "https://www.google.it", List<string> extensions = null)
        {
            browser = browser_;
            switch (browser)
            {
                case "CR":
                    chrome = ApriBrowser_CR(link, extensions);
                    return chrome;

                case "FF":
                    firefox = ApriBrowser_FF(link);
                    if (extensions != null)
                    {
                        foreach (string extension in extensions)
                        {
                            firefox.InstallAddOnFromFile(@"K:\SebiAutomation\Resources\Extensions\" + extension);
                        }
                    }
                    return firefox;

                case "ED":
                    edge = ApriBrowser_ED(link, extensions);
                    return edge;
            }

            return null;
        }
        public static ChromeDriver ApriBrowser_CR(string link = "https://www.google.it", List<string> extensions = null)
        {
            browser = "CR";
            ChromeDriverService chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            ChromeOptions opt = new ChromeOptions();
            if (extensions != null)
            {
                foreach (string ext in extensions)
                {
                    opt.AddExtension(@"K:\SebiAutomation\Resources\Extensions\" + ext);
                }
            }

            if (!automate)
            {
                Process.Start("chrome.exe", " --remote-debugging-port=9222 --disk-cache-size=1 --media-cache-size=1");
                opt.DebuggerAddress = "localhost:9222";
            }
            else
            {
                opt.AddUserProfilePreference("download.default_directory", @"C:\A\");
            }
            opt.AddArgument("--disk-cache-size=1");
            opt.AddArgument("--media-cache-size=1");
            opt.AddArgument("--disable-search-engine-choice-screen");

            chrome = new ChromeDriver(chromeDriverService, opt, TimeSpan.FromMinutes(5));

            chrome.Manage().Window.Maximize();
            chrome.Navigate().GoToUrl(link);

            return chrome;
        }
        public static FirefoxDriver ApriBrowser_FF(string link)
        {
            browser = "FF";

            var firefoxDriverService = FirefoxDriverService.CreateDefaultService();
            firefoxDriverService.HideCommandPromptWindow = true;

            var opt = new FirefoxOptions();
            firefox = new FirefoxDriver(firefoxDriverService, opt, TimeSpan.FromMinutes(5));

            firefox.Manage().Window.Maximize();
            firefox.Navigate().GoToUrl(link);

            return firefox;
        }
        public static EdgeDriver ApriBrowser_ED(string link, List<string> extensions = null)
        {
            browser = "ED";

            var edgeDriverService = EdgeDriverService.CreateDefaultService();
            edgeDriverService.HideCommandPromptWindow = true;

            var opt = new EdgeOptions();

            if (extensions != null)
            {
                foreach (string ext in extensions)
                {
                    opt.AddExtension(@"K:\SebiAutomation\Resources\Extensions\" + ext);
                }
            }

            if (!automate)
            {
                Process.Start("msedge.exe", " --remote-debugging-port=9229 --disk-cache-size=1 --media-cache-size=1");
                opt.DebuggerAddress = "localhost:9229";
            }
            else
            {
                opt.AddUserProfilePreference("download.default_directory", @"C:\A\");
            }

            edge = new EdgeDriver(edgeDriverService, opt, TimeSpan.FromMinutes(5));
            edge.Manage().Window.Maximize();
            edge.Navigate().GoToUrl(link);

            return edge;
        }
        public static string Riapri(string link = "")
        {
            string handler = string.Empty;
            string new_handler = string.Empty;
            System.Drawing.Point pos;


            switch (browser)
            {
                case "CR":
                    pos = chrome.Manage().Window.Position;
                    handler = chrome.CurrentWindowHandle;
                    chrome.SwitchTo().Window(handler);
                    chrome.Close();

                    clean();

                    chrome.SwitchTo().NewWindow(WindowType.Window);
                    new_handler = chrome.CurrentWindowHandle;

                    chrome.SwitchTo().Window(new_handler);
                    chrome.Navigate().GoToUrl(link);

                    chrome.ExecuteJavaScript("window.open('https://www.w3schools.com')");

                    chrome.Manage().Window.Position = pos;
                    chrome.Manage().Window.Maximize();

                    break;

                case "FF":
                    pos = firefox.Manage().Window.Position;
                    handler = firefox.CurrentWindowHandle;

                    firefox.SwitchTo().NewWindow(WindowType.Window);
                    new_handler = firefox.CurrentWindowHandle;

                    firefox.SwitchTo().Window(handler);
                    firefox.Close();
                    firefox.SwitchTo().Window(new_handler);
                    firefox.Navigate().GoToUrl(link);
                    firefox.Manage().Window.Maximize();
                    break;

                case "ED":
                    pos = edge.Manage().Window.Position;
                    handler = edge.CurrentWindowHandle;

                    edge.SwitchTo().NewWindow(WindowType.Window);
                    new_handler = edge.CurrentWindowHandle;

                    edge.SwitchTo().Window(handler);
                    edge.Close();
                    clean();
                    edge.SwitchTo().Window(new_handler);
                    edge.Navigate().GoToUrl(link);

                    edge.Manage().Window.Position = pos;
                    edge.Manage().Window.Maximize();
                    break;
            }
            return new_handler;
        }
        public static string Page()
        {
            string ret = "";

            switch (browser)
            {
                case "CR":
                    ret = chrome.Title;
                    break;

                case "FF":
                    ret = firefox.Title;
                    break;

                case "ED":
                    ret = edge.Title;
                    break;
            }

            return ret;
        }
        public static bool vedisece(string cosa, IWebDriver contesto = null, int timer = 10)
        {
            if (contesto == null)
            {
                switch (browser)
                {
                    case "CR":
                        contesto = chrome;
                        break;

                    case "FF":
                        contesto = firefox;
                        break;

                    case "ED":
                        contesto = edge;
                        break;

                    default:
                        browser = "CR";
                        contesto = chrome;
                        break;
                }
            }

            WebDriverWait wait = new WebDriverWait(contesto, TimeSpan.FromSeconds(timer));
            string x = wait.Until(d => d.FindElement(By.TagName("body"))).GetDomProperty("innerHTML");
            if (x.Contains(cosa))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Cattura la tabella HTML tramite una stringa xpath
        /// </summary>
        /// <param name="attributo">Attributo identificatore per XPath es. [@class='listaIsp4']</param>
        /// <param name="contesto">Contesto in cui cercare la tabella, default in tutto il documento</param>
        /// <returns></returns>
        public static List<List<string>> CaptureTable(string attributo = "", IWebElement contesto = null)
        {
            IList<IWebElement> elementi;
            List<string> lista_riga = new List<string>();
            List<List<string>> result = new List<List<string>>();
            String xpath = string.Empty;
            String tmp = string.Empty;
            MessageFrame msg = new MessageFrame(4);

            try
            {
                int ricerca = 1;
                int r = 1;

                while (true)
                {
                    msg.scrivi("Catturo riga tabella " + r);

                    if (attributo != "")
                    {
                        xpath = "//table" + attributo;
                    }
                    else
                    {
                        xpath = "//table";
                    }
                    xpath += "/tbody/tr[" + r + "]/td";

                    if (contesto != null) // ricerca in elemento
                    {
                        elementi = contesto.FindElements(By.XPath("." + xpath));
                    }
                    else // ricerca in intero documento
                    {
                        elementi = chrome.FindElements(By.XPath(xpath));
                    }
                    if (elementi.Count == 0)
                    {
                        r += 1;
                        ricerca += 1;
                        if (ricerca > 2)
                        {
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    foreach (IWebElement elemento in elementi)
                    {
                        tmp = elemento.GetDomProperty("innerHTML");
                        //tmp = elemento.Text;
                        tmp = tmp.Replace("<br>", "");
                        tmp = tmp.Replace("<sup>", "");
                        tmp = tmp.Replace("</sup>", "");
                        tmp = tmp.Replace("&amp;", "&");
                        tmp = tmp.Trim();

                        lista_riga.Add(tmp);
                    }
                    result.Add(lista_riga);
                    lista_riga = new List<string>();
                    r += 1;
                }
                msg.chiudi();
            }
            catch (Exception)
            {
            }
            return result;
        }
        public static int download(string destPath, string titolo_file = "", string formato_file = ".*", string file_scaricato = "")
        {
            Thread.Sleep(1000);

            int tentativo = 0;
            string sourcePath = @"C:\A\";
            bool trovato = false;
            string[] sourceFile = Directory.GetFiles(sourcePath, "*.*");

            if (sourceFile.Length > 1)
            {
                foreach (string f in sourceFile)
                {
                    File.Delete(f);
                }
                sourceFile = Directory.GetFiles(sourcePath, "*.*");
                if (sourceFile.Length > 1)
                {
                    MsgBox.Show("Attenzione! Più di un file in C:\\A\\, eliminare tutti i file prima di continuare");
                    throw new Exception("Attenzione ! Più di un file in C:\\A\\");
                }
                return -3;
            }
            else if (sourceFile.Length == 0)
            {
                for (tentativo = 0; tentativo < 7; tentativo++)
                {
                    Thread.Sleep(5000);
                    sourceFile = Directory.GetFiles(sourcePath, "*" + formato_file);
                    if (sourceFile.Length == 1)
                    {
                        trovato = true;
                        break;
                    }
                }
            }
            else
            {
                if (file_scaricato != "")
                {
                    for (tentativo = 0; tentativo < 10; tentativo++)
                    {
                        if (Utils.desa(sourceFile[0], @"C:\A\") == file_scaricato | sourceFile[0].Contains("EPA"))
                        {
                            trovato = true;
                            break;
                        }
                        else
                        {
                            Thread.Sleep(1000);
                            sourceFile = Directory.GetFiles(sourcePath, "*" + formato_file);
                            continue;
                        }
                    }
                    if (!trovato)
                    {
                        MsgBox.Show("ATTENZIONE!! Non sono riuscito a scaricare il file, controllare la cartella C:\\A\\");
                        Environment.Exit(1);
                    }
                }
                else
                {
                    if (titolo_file == "")
                    {
                        titolo_file = Utils.ExtractBetween(sourceFile[0], sourcePath, ".");
                    }
                    trovato = true;
                }
            }

            if (trovato)
            {
                if (formato_file == ".*")
                {
                    formato_file = sourceFile[0].Substring(sourceFile[0].IndexOf('.'));
                }
                titolo_file += formato_file;

                tentativo = 0;
                while (true)
                {
                    try
                    {
                        File.Move(sourceFile[0], destPath + titolo_file);
                        Thread.Sleep(2000);
                        if (!File.Exists(destPath + titolo_file))
                        {
                            MsgBox.Show("C'è qualcosa che non va, file " + titolo_file + " non spostato in " + destPath + ", contattare reparto tecnico");
                            Environment.Exit(1);
                        }
                        if (File.Exists(sourceFile[0]))
                        {
                            MsgBox.Show("C'è qualcosa che non va, file " + sourceFile[0] + " ancora in cartella , contattare reparto tecnico");
                            Environment.Exit(1);
                        }
                        return 0;
                    }
                    catch (IOException ex)
                    {
                        if (ex.Message.Contains("esiste"))
                        {
                            File.Delete(sourceFile[0]);
                            return -2;
                        }
                        else if (ex.Message.Contains("non è stato trovato"))
                        {
                            Thread.Sleep(1000);
                            sourceFile = Directory.GetFiles(sourcePath, "*" + formato_file);
                            continue;
                        }
                        else
                        {
                            tentativo += 1;
                            if (tentativo < 10)
                            {
                                Thread.Sleep(500);
                                continue;
                            }
                            else
                            {
                                break;
                            }
                        }
                        //err.GestiscoErr(sourceFile[0] + " non spostato perchè già presente in " + destPath + " - CONSULTA ERROR.LOG", 3, ex);
                        //File.Delete(sourceFile[0]);
                    }
                }
            }
            return -1;
        }
        public static int SelectHTML(IWebDriver contesto, string tipoElemento, string identificativo, string cosa, int timeout = 10)
        {
            WebDriverWait wait = new WebDriverWait(contesto, TimeSpan.FromSeconds(timeout));
            try
            {
                switch (tipoElemento)
                {
                    case "link":
                        wait.Until(d => d.FindElement(By.LinkText(identificativo))).SendKeys(cosa);
                        break;
                    case "id":
                        wait.Until(d => d.FindElement(By.Id(identificativo))).SendKeys(cosa);
                        break;
                }
            }
            catch (WebDriverTimeoutException)
            {
                return -1;
            }
            return 0;
        }
        public static void chiudiDriver()
        {
            List<Process> proc = new List<Process>();
            switch (browser)
            {
                case "CR":
                    proc = Process.GetProcessesByName("chromedriver").ToList();
                    break;

                case "FF":
                    proc = Process.GetProcessesByName("geckodriver").ToList();
                    break;

                case "ED":
                    proc = Process.GetProcessesByName("msedgedriver").ToList();
                    break;
            }

            foreach (Process p in proc)
            {
                try
                {
                    p.Kill();
                }
                catch { }
            }
        }
        public static void clean()
        {
            string percorso = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Google\Chrome\User Data\Default\";
            List<string> cartelle = new List<string>() { "History", "Cache\\Cache_data\\", "Network\\Cookies", "Local Storage\\leveldb\\", "Session Storage\\" };


            switch (browser)
            {
                case "CR":

                    try
                    {
                        foreach (string cartella in cartelle)
                        {
                            string param = "*";
                            string cartella_ = cartella;

                            if (!cartella.EndsWith("\\"))
                            {
                                string tmp = cartella;

                                if (tmp.Contains("\\"))
                                {
                                    tmp = Utils.desa(cartella, "\\");
                                }

                                param = tmp + "*";

                                cartella_ = cartella_.Replace(tmp, "");
                            }

                            string[] files = Directory.GetFiles(percorso + cartella_, param);

                            foreach (string file in files)
                            {
                                while (true)
                                {
                                    try
                                    {
                                        File.Delete(file);
                                        break;
                                    }
                                    catch { continue; }
                                }
                            }
                        }

                    }
                    catch { }

                    break;

                case "ED":

                    try
                    {
                        foreach (string cartella in cartelle)
                        {
                            string param = "*";
                            string cartella_ = cartella;

                            if (!cartella.EndsWith("\\"))
                            {
                                string tmp = cartella;

                                if (tmp.Contains("\\"))
                                {
                                    tmp = Utils.desa(cartella, "\\");
                                }

                                param = tmp + "*";

                                cartella_ = cartella_.Replace(tmp, "");
                            }

                            string[] files = Directory.GetFiles(percorso + cartella_, param);

                            foreach (string file in files)
                            {
                                while (true)
                                {
                                    try
                                    {
                                        File.Delete(file);
                                        break;
                                    }
                                    catch { continue; }
                                }
                            }
                        }

                    }
                    catch { }

                    break;

                case "FF":

                    break;

            }
        }
    }
    public class LogFile
    {
        private LoggerConfiguration logOptions = new LoggerConfiguration();
        public readonly string filename;
        private Serilog.Core.Logger logger;
        public LogFile(string nome_file = null, bool logTemplate = false)
        {
            if (nome_file == null)
            {
                this.filename = Process.GetCurrentProcess().ProcessName
                    + DateTime.Now.ToString("_yyyyMMdd_HHmm") + ".log";
                logTemplate = true;
            }
            else
            {
                filename = nome_file;
            }

            if (logTemplate)
            {
                logOptions.MinimumLevel.Debug();
                logOptions.WriteTo.File(filename,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] || {Message}{NewLine}");
            }
            else
            {
                logOptions.WriteTo.File(filename,
                    outputTemplate: "{Message}{NewLine}");
            }

            logger = logOptions.CreateLogger();
        }
        public void Write(string messaggio, int gravita = 0, Exception ex = null)
        {
            switch (gravita)
            {
                case 1:
                    logger.Debug(messaggio);
                    break;

                case 2:
                    logger.Warning(messaggio);
                    break;

                case 3:
                    if (ex != null)
                    {
                        LogFile error = new LogFile("ERROR.LOG");
                        error.Write("======================================================================================");
                        error.Write("ERRORE DEL " + DateTime.Now.ToString("dd/MM/yyyy - HH:mm:ss"));
                        error.Write("======================================================================================");
                        error.Write(ex.ToString());
                        error.Chiudi();

                        logger.Error("CONTROLLA FILE 'ERROR.LOG'");
                    }

                    logger.Error(messaggio);
                    break;

                case 4:
                    if (ex != null)
                    {
                        var fatal = new LogFile("FATALERR.LOG");
                        fatal.Write("======================================================================================");
                        fatal.Write("ERRORE DEL " + DateTime.Now.ToString("dd/MM/yyyy - hh:mm:ss"));
                        fatal.Write("======================================================================================");
                        fatal.Write(ex.ToString());
                        fatal.Chiudi();

                        MsgBox.Show(String.Format("{0}.\nControllare 'FATALERR.LOG'", ex.Message), "FATAL", MessageBoxImage.Error);

                        logger.Fatal("CONTROLLA FILE 'FATALERR.LOG'");
                    }

                    logger.Fatal(messaggio);
                    Chiudi();
                    Environment.Exit(0);
                    break;

                case -1:
                    logger.Verbose(messaggio);
                    break;

                default:
                    logger.Information(messaggio);
                    break;
            }
        }
        public void Chiudi()
        {
            logger.Dispose();
        }
    }
    public static class MsgBox
    {
        public static MessageBoxResult Show(string messaggio, string titolo = "MsgBox", MessageBoxImage icona = MessageBoxImage.Information, MessageBoxButton pulsanti = MessageBoxButton.OK)
        {
            MessageBoxResult ret = MessageBox.Show(messaggio, titolo, pulsanti, icona, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);
            return ret;
        }
    }
    public static class Ivass
    {
        public static bool societa = false;

        public static ChromeDriver driver = null;
        static WebDriverWait wait;
        public static List<List<string>> cerca()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            List<List<string>> result = new List<List<string>>();

            try
            {
                wait.Until(d => d.FindElement(By.XPath("//*[@id=\"compagnia\"]/div/div/form/div[1]/div[2]/label"))).Click();
                Thread.Sleep(500);

                new Actions(driver).SendKeys(Keys.End).Perform();
                Thread.Sleep(1000);

                while (true)
                {
                    try
                    {
                        wait.Until(d => d.FindElement(By.XPath("//*[@id='compagnia']/div/div/form/div[3]/div/button[3]"))).Click(); // CLICK RICERCA
                        break;
                    }
                    catch { new Actions(driver).SendKeys(Keys.End).Perform(); continue; }
                }
            }
            catch { }

            Thread.Sleep(3000);
            aspetta_caricamento();

            int tentativi = 0;
            while (true)
            {
                if (driver.Title == "Dati Registro")
                {
                    result = HTML.CaptureTable("[@class='table table-gest table-striped']");

                    break;
                }
                else
                {
                    if (HTML.vedisece("Non esistono intermediari per la selezione effettuata"))
                    {
                        break;
                    }

                    if (tentativi > 30)
                    {
                        break;
                    }
                    else
                    {
                        try
                        {
                            wait.Until(d => d.FindElement(By.XPath("//*[@id='compagnia']/div/div/form/div[3]/div/button[3]"))).Click(); // CLICK RICERCA
                        }
                        catch { new Actions(driver).SendKeys(Keys.End).Perform(); }
                        aspetta_caricamento();

                        tentativi += 1;

                        continue;
                    }
                }
            }
            return result;
        }
        public static Dictionary<string, string> ricerca(string numero_iscr, string confronta = "", string nome = "", string cognome = "")
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            List<List<string>> quanti_result = new List<List<string>>();
            Dictionary<string, string> dettaglio = new Dictionary<string, string>();
            bool ricerca_sogg = true;
            bool ricerca_soc = false;

            bool ricerca_niscr = false;
            bool ricerca_denom = false;

            Thread.Sleep(1000);

            while (true)
            {
                aspetta_caricamento();

                if (numero_iscr != "" && !ricerca_niscr)
                {
                    wait.Until(d => d.FindElement(By.Id("matricola"))).SendKeys(numero_iscr);
                    Thread.Sleep(500);
                    quanti_result = cerca();
                    ricerca_niscr = true;
                }
                else
                {
                    if (cognome != "") // SE COGNOME E' VALORIZZATO TRATTASI DI SOGGETTO FISICO
                    {
                        ricerca_niscr = false;
                        ricerca_denom = true;

                        ricerca_soc = false;

                        wait.Until(d => d.FindElement(By.Id("cognome"))).SendKeys(cognome);
                        Thread.Sleep(500);
                        wait.Until(d => d.FindElement(By.Id("nome"))).SendKeys(nome);
                        quanti_result = cerca();
                    }
                    else // TRATTASI DI SOCIETA'
                    {
                        ricerca_soc = true;
                        Thread.Sleep(1000);

                        wait.Until(d => d.FindElement(By.XPath("//input[@value='PG']"))).Click();
                        Thread.Sleep(500);
                        wait.Until(d => d.FindElement(By.Id("ragSoc"))).SendKeys(nome);
                        Thread.Sleep(500);
                        quanti_result = cerca();
                        ricerca_denom = true;
                    }
                }

                if (quanti_result.Count > 0) // SE TROVATI RISULTATI
                {
                    while (true)
                    {
                        string x = wait.Until(d => d.FindElement(By.XPath("//li[(contains(@class, 'page-item')) and (contains(@class, 'active'))]/a"))).GetAttribute("innerHTML");

                        int active_pag = int.Parse(x.Substring(0, x.IndexOf("<span")).Trim());
                        IList<IWebElement> tmp = wait.Until(d => d.FindElements(By.XPath("//a[contains(@class, 'page-link')]")));

                        string s = tmp[tmp.Count - 2].GetDomProperty("innerHTML");
                        int pagine = 0;
                        if (s.Contains("<span"))
                        {
                            pagine = int.Parse(s.Substring(0, s.IndexOf("<span")).Trim());
                        }
                        else if (s.Contains("<!---->"))
                        {
                            pagine = int.Parse(s.Replace("<!---->", ""));
                        }
                        else
                        {
                            pagine = int.Parse(tmp[tmp.Count - 2].GetAttribute("innerHTML").Trim());
                        }

                        IList<IWebElement> element_result = wait.Until(d => d.FindElements(By.XPath("//table[@class='table table-gest table-striped']/tbody/tr")));

                        for (int i = 0; i < element_result.Count; i++)
                        {

                            aspetta_caricamento();

                            Thread.Sleep(3000);
                            if (ricerca_soc) // SE HO RICERCATO PER SOCIETA'
                            {
                                IWebElement elem = element_result[i];
                                elem.Click();
                                dettaglio = cattura_dettagli();

                                if (confronta != "" && numero_iscr == "")
                                { // CONFRONTO INDIRIZZO
                                    dettaglio["SEDE LEGALE"] = Utils.TogliAccenti(dettaglio["SEDE LEGALE"]);
                                    if (dettaglio["SEDE LEGALE"].Contains(confronta))
                                    {
                                        new Actions(driver).SendKeys(Keys.End).Perform();
                                        Thread.Sleep(500);
                                        while (true)
                                        {
                                            try
                                            {
                                                wait.Until(d => d.FindElement(By.XPath("//div[@class='btn-box']/div[2]/button"))).Click();
                                                Thread.Sleep(500);
                                                break;
                                            }
                                            catch (ElementClickInterceptedException)
                                            {
                                                new Actions(driver).SendKeys(Keys.End).Perform();
                                                Thread.Sleep(500);
                                                continue;
                                            }
                                        }

                                        return dettaglio;
                                    }
                                }
                                else if (ricerca_niscr || dettaglio["RAGIONE O DENOMINAZIONE SOCIALE"].Trim().Contains(nome))
                                {
                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                    Thread.Sleep(500);
                                    while (true)
                                    {
                                        try
                                        {
                                            wait.Until(d => d.FindElement(By.XPath("//div[@class='btn-box']/div[2]/button"))).Click();
                                            Thread.Sleep(500);
                                            break;
                                        }
                                        catch (ElementClickInterceptedException)
                                        {
                                            new Actions(driver).SendKeys(Keys.End).Perform();
                                            Thread.Sleep(500);
                                            continue;
                                        }
                                    }

                                    return dettaglio;
                                }
                            }
                            else // ALTRIMENTI HO RICERCATO PER SOGGETTO
                            {
                                IWebElement elem = element_result[i];
                                bool successivo = false;

                                if (confronta != "" & (numero_iscr == "" | !ricerca_niscr))
                                { // CONFRONTO DATA DI NASCITA
                                    while (true)
                                    {
                                        try
                                        {
                                            if (elem.Text.Contains(" " + cognome + " "))
                                            {
                                                elem.Click();
                                                dettaglio = cattura_dettagli();
                                            }
                                            else
                                            {
                                                successivo = true;
                                            }
                                            break;
                                        }
                                        catch (Exception ex)
                                        {
                                            new Actions(driver).SendKeys(Keys.End).Perform();
                                        }
                                    }

                                    if (successivo)
                                    {
                                        continue;
                                    }

                                    if (confronta == dettaglio["DATA NASCITA"])
                                    {
                                        try
                                        {
                                            wait.Until(d => d.FindElement(By.XPath("//*[@id='anagrafica']/div/div[1]/div[1]/label"))).Click(); // CLICK FUOCO
                                        }
                                        catch { }

                                        new Actions(driver).SendKeys(Keys.End).Perform();
                                        Thread.Sleep(500);
                                        while (true)
                                        {
                                            try
                                            {
                                                wait.Until(d => d.FindElement(By.XPath("//div[@class='btn-box']/div[2]/button"))).Click(); // NUOVA RICERCA
                                                Thread.Sleep(500);
                                                break;
                                            }
                                            catch (ElementClickInterceptedException)
                                            {
                                                new Actions(driver).SendKeys(Keys.End).Perform();
                                                Thread.Sleep(500);
                                                continue;
                                            }
                                        }
                                        return dettaglio;
                                    }
                                }
                                else if (ricerca_niscr)
                                {
                                    elem.Click();
                                    dettaglio = cattura_dettagli();

                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                    Thread.Sleep(500);
                                    while (true)
                                    {
                                        try
                                        {
                                            wait.Until(d => d.FindElement(By.XPath("//div[@class='btn-box']/div[2]/button"))).Click(); // NUOVA RICERCA
                                            Thread.Sleep(500);
                                            break;
                                        }
                                        catch (ElementClickInterceptedException)
                                        {
                                            new Actions(driver).SendKeys(Keys.End).Perform();
                                            Thread.Sleep(500);
                                            continue;
                                        }
                                    }
                                    return dettaglio;
                                }
                            }

                            while (true)
                            {
                                try
                                {
                                    wait.Until(d => d.FindElement(By.XPath("//div[@class='btn-box']/div/button"))).Click(); // INDIETRO

                                    break;
                                }
                                catch (Exception ex)
                                {
                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                }
                            }
                            Thread.Sleep(2000);

                            aspetta_caricamento();

                            for (int pag = 1; pag < active_pag; pag++)
                            {
                                try
                                {

                                    tmp = wait.Until(d => d.FindElements(By.XPath("//ul[@class='pagination']/li"))); // ANDREA 
                                    //wait.Until(d => d.FindElement(By.XPath("//*[@id='sub-navbar']/elenco-registro-unico-intermediari/div/h3"))).Click(); // CLICK FUOCO
                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                    tmp[tmp.Count - 1].Click();
                                }
                                catch
                                {
                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                    pag -= 1;
                                    continue;
                                }

                                Thread.Sleep(2000);
                                aspetta_caricamento();

                                x = wait.Until(d => d.FindElement(By.XPath("//li[(contains(@class, 'page-item')) and (contains(@class, 'active'))]/a"))).GetDomProperty("innerHTML");
                                int cur_page = int.Parse(x.Substring(0, x.IndexOf("<span")).Trim());
                                if ((pag + 1) < cur_page)
                                {
                                    pag -= 1;
                                }
                            }

                            Thread.Sleep(3000);
                            element_result = wait.Until(d => d.FindElements(By.XPath("//table[@class='table table-gest table-striped']/tbody/tr")));
                        }

                        if (pagine > active_pag)
                        {
                            while (true)
                            {
                                try
                                {
                                    tmp = wait.Until(d => d.FindElements(By.XPath("//ul[@class='pagination']/li"))); // ANDREA
                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                    tmp[tmp.Count - 1].Click();
                                    Thread.Sleep(2000);

                                    aspetta_caricamento();
                                    break;
                                }
                                catch
                                {
                                    new Actions(driver).SendKeys(Keys.End).Perform();
                                    continue;
                                }
                            }

                            quanti_result = HTML.CaptureTable("[@class='table table-gest table-striped']");
                            continue;
                        }
                        else
                        {
                            break;
                        }
                    }

                    while (true)
                    {
                        try
                        {
                            wait.Until(d => d.FindElement(By.XPath("//elenco-registro-unico-intermediari/div/div/div[3]/button[1]"))).Click();  // NUOVA RICERCA

                            break;
                        }
                        catch (Exception ex)
                        {
                            new Actions(driver).SendKeys(Keys.End).Perform();
                            continue;
                        }
                    }
                    Thread.Sleep(500);
                    if (ricerca_denom)
                    {
                        return new Dictionary<string, string>();
                    }
                }
                else
                {
                    if (HTML.vedisece("Accesso negato", driver))
                    {
                        driver.Navigate().GoToUrl("https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                        Thread.Sleep(500);
                        wait.Until(d => d.FindElement(By.XPath("//div[@class='btn-box']/div/button"))).Click(); // INDIETRO
                        Thread.Sleep(500);

                        ricerca_niscr = false;
                        ricerca_denom = false;
                        continue;
                    }

                    if (ricerca_sogg)
                    {
                        if (!ricerca_soc && numero_iscr != ""/* && cognome == "" */)
                        {
                            ricerca_soc = true;
                            ricerca_niscr = false;
                            societa = true;
                            try
                            {
                                wait.Until(d => d.FindElement(By.XPath("//elenco-registro-unico-intermediari/div/div/div[3]/button[1]"))).Click(); // NUOVA RICERCA
                            }
                            catch { }

                            while (true)
                            {
                                try
                                {
                                    Thread.Sleep(1000);
                                    wait.Until(d => d.FindElement(By.XPath("//input[@value='PG']"))).Click();
                                    break;
                                }
                                catch { continue; }
                            }

                            Thread.Sleep(500);
                            continue;
                        }
                        else
                        {
                            if (HTML.vedisece("Accesso negato", driver))
                            {
                                driver.Navigate().GoToUrl("https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                                Thread.Sleep(500);
                                wait.Until(d => d.FindElement(By.XPath("//elenco-registro-unico-intermediari/div/div/div[3]/button[1]"))).Click();  // NUOVA RICERCA

                                Thread.Sleep(500);
                                continue;
                            }
                            else
                            {
                                wait.Until(d => d.FindElement(By.XPath("//elenco-registro-unico-intermediari/div/div/div[3]/button[1]"))).Click();  // NUOVA RICERCA

                                if ((ricerca_denom || nome == "") && (ricerca_niscr || numero_iscr == ""))
                                {
                                    // NESSUN RISULTATO
                                    return new Dictionary<string, string>();
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
        }
        public static Dictionary<string, string> cattura_dettagli()
        {
            //Thread.Sleep(2000);
            aspetta_caricamento();
            Dictionary<string, string> result = new Dictionary<string, string>();
            List<IWebElement> dettagli = new List<IWebElement>(wait.Until(d => d.FindElements(By.XPath("//div[@id='anagrafica']/div/div"))));  //ANDREA

            foreach (IWebElement element in dettagli)
            {
                string descrizione = element.Text;
                string valore = Utils.ExtractBetween(descrizione, "\r\n", "");
                descrizione = Utils.ExtractBetween(descrizione, "", "\r\n");
                if (descrizione == "")
                {
                    continue;
                }
                valore = valore.Replace(descrizione, "");

                if (descrizione != "")
                {
                    result.Add(descrizione.ToUpper(), valore);
                }
            }

            List<IWebElement> elem = new List<IWebElement>(wait.Until(d => d.FindElements(By.XPath("//div[@id='titoloIndividuale']/div/div"))));
            foreach (IWebElement element in elem)
            {
                string descrizione = element.Text;
                string valore = Utils.ExtractBetween(descrizione, "\r\n", "");
                descrizione = Utils.ExtractBetween(descrizione, "", "\r\n");

                valore = valore.Replace(descrizione, "");

                if (descrizione != "")
                {
                    result.Add(descrizione, valore);
                }
            }

            if (result.ContainsKey("Cariche societarie"))
            {
                elem = new List<IWebElement>(wait.Until(d => d.FindElements(By.XPath("//div[@id='titoloSocietario']/div/div"))));

                foreach (IWebElement element in elem)
                {
                    string descrizione = element.Text;
                    string valore = Utils.ExtractBetween(descrizione, "\r\n", "");
                    descrizione = Utils.ExtractBetween(descrizione, "", "\r\n");

                    valore = valore.Replace(descrizione, "");
                    if (valore.Contains("Societa"))
                    {
                        valore = valore.Replace("Societa", "");
                    }

                    if (descrizione != "")
                    {
                        result.Add(descrizione, valore);
                    }
                }
            }

            return result;
        }
        public static void collega_driver(ChromeDriver browser)
        {
            driver = browser;
            HTML.chrome = browser;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
        public static void aspetta_caricamento()
        {
            while (true)
            {
                if (!wait.Until(d => d.FindElement(By.XPath("//app-loader/div"))).Displayed)
                {
                    break;
                }
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
        }
    }
    public class PDF
    {
        public XFont font { get; set; } = new XFont("Tahoma", 12);
        public PdfDocument pdfDocument { get; set; }
        public PdfPage pdfPage { get; set; }
        public XGraphics gfx { get; set; }
        public XTextFormatter tf { get; set; }

        public int pagina = 1;
        private XUnit margin_y = XUnit.FromPoint(50);

        public XUnit pos_y = XUnit.FromPoint(50);
        public XUnit pos_x = XUnit.FromPoint(60);
        public XUnit spacing = XUnit.FromPoint(7);

        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Temp\\tmp.pdf";
        /// <summary>
        /// Crea documento PDF
        /// </summary>
        /// <param name="percorso"> Se percorso = null userà come percorso AppData\Local\Temp  </param>
        public PDF(string percorso = null)
        {
            percorso = percorso == null ? documentsPath : percorso;

            //if (!Directory.Exists(percorso))
            //{
            //    Directory.CreateDirectory(percorso);
            //}

            this.pdfDocument = new PdfDocument();
            this.pdfPage = pdfDocument.AddPage();

            this.gfx = XGraphics.FromPdfPage(pdfPage);
            this.tf = new XTextFormatter(gfx);
            XFont font = new XFont("Tahoma", 12);

            this.header("Pag. " + pagina, font, null, 2, 150);
        }
        /// <summary>
        /// Scrive stringa in pagina PDF
        /// </summary>
        /// <param name="cosa">Stringa da scrivere, per il ritorno a capo inserire l'escape '\n' seguito da uno spazio</param>
        /// <param name="colore"></param>
        /// <param name="font"></param>
        /// <param name="add_y">Riempi per incrementare la posizione Y iniziale</param>
        /// <param name="pos_x">Posizione X iniziale, default = 50</param>
        public void scrivi(string cosa, XSolidBrush colore, bool centered = false, XFont font = null, int add_y = 0, double pos_x = 0)
        {
            if (pos_x >= 0 - Costanti.epsilon || pos_x <= 0 + Costanti.epsilon)
            {
                pos_x = this.pos_x.Point;
            }

            List<string> split_cosa = cosa.Split(' ').ToList();
            while (split_cosa.Count > 0)
            {
                split_cosa = this.scrivi_riga(split_cosa, colore, centered, font, add_y, pos_x);
                if (this.pos_y >= this.pdfPage.Height - (this.margin_y * 2))
                {
                    this.addPage();
                }
            }
        }
        public List<string> scrivi_riga(List<string> split_cosa, XSolidBrush colore, bool centered = false, XFont font = null, int add_y = 0, double pos_x = 60)
        {
            this.tf.Font = font == null ? this.font : font;
            //if (add_y > 0)
            if (add_y > 0)
            {
                this.pos_y += XUnit.FromPoint(add_y);
            }

            XStringFormat string_format = XStringFormats.TopLeft;
            if (centered)
            {
                string_format = XStringFormats.Center;
            }

            string riga = "";
            bool scrivi = false;
            while (true)
            {
                XSize dim = this.gfx.MeasureString(riga, this.tf.Font, string_format);
                if (XUnit.FromPoint(dim.Width) < (this.pdfPage.Width - (this.pos_x * 2)) & !scrivi)
                {
                    riga += split_cosa[0];
                    split_cosa.RemoveAt(0);

                    if (riga.Contains("\n") & !riga.StartsWith("\n"))
                    {
                        List<string> y = new List<string>();
                        string tmp = string.Empty;
                        foreach (char c in riga)
                        {
                            tmp += c.ToString();
                            if (tmp.Contains("\n"))
                            {
                                y.Add(tmp);
                                tmp = string.Empty;
                            }
                        }
                        if (tmp != string.Empty)
                        {
                            y.Add(tmp);
                            tmp = string.Empty;
                        }

                        for (int i = y.Count - 1; i > 0; i--)
                        {
                            split_cosa.Insert(0, y[i]);
                        }
                        riga = y[0] + "\n";
                    }

                    dim = this.gfx.MeasureString(riga, this.tf.Font, string_format);
                    if (riga.Contains("\n") | split_cosa.Count == 0 |
                                    XUnit.FromPoint(this.gfx.MeasureString(riga + (split_cosa.Count > 0 ? split_cosa[0] : ""),
                                    this.tf.Font, string_format).Width) > this.pdfPage.Width - this.pos_x - XUnit.FromPoint(pos_x))
                    {
                        riga = riga.Replace("\t", "  ");
                        scrivi = true;
                    }
                    riga += " ";
                }
                else
                {
                    gfx.DrawString(riga.TrimEnd(), this.tf.Font, colore, new XRect(pos_x, this.pos_y.Point, (this.pdfPage.Width - (this.pos_x * 2)).Point, dim.Height), string_format);
                    break;
                }
            }

            if (add_y >= 0)
            {
                this.pos_y += XUnit.FromPoint(this.font.Height) + this.spacing;
            }
            return split_cosa;
        }
        public void set_font(string font, int size, XFontStyleEx stile_font = XFontStyleEx.Regular)
        {
            this.font = new XFont(font, size, stile_font);
        }
        public void addY(int add_y)
        {
            this.pos_y += XUnit.FromPoint(add_y);

            if (this.pos_y > this.pdfPage.Height - XUnit.FromPoint(70))
            {
                addPage();
            }
        }
        /// <summary>
        /// Inserisce un immagine come intestazione del PDF
        /// </summary>
        /// <param name="pos">Se 0 posiziona a SX, se 1 posiziona al Centro, se 2 posiziona DX</param>
        public void header(XImage immagine, int pos = 0)
        {
            XUnit x = XUnit.FromPoint(0);
            double pos_y = 30 - (immagine.PointHeight / 5);
            switch (pos)
            {
                case 0:
                    x = XUnit.FromPoint(50);
                    break;

                case 1:
                    x = (this.pdfPage.Width / 2) - XUnit.FromPoint(immagine.PointWidth / 2);
                    break;

                case 2:
                    x = (this.pdfPage.Width - XUnit.FromPoint(50)) - XUnit.FromPoint(immagine.PointWidth);
                    break;
            }
            this.gfx.DrawImage(immagine, x.Point, pos_y);

            this.pos_y += XUnit.FromPoint(20);
        }
        /// <summary>
        /// Inserisce una stringa come intestazione del PDF
        /// </summary>
        /// <param name="cosa">Stringa da scrivere, per il ritorno a capo inserire l'escape '\n' seguito da uno spazio</param>
        /// <param name="pos">Se 0 posiziona a SX, se 1 posiziona al Centro, se 2 posiziona DX</param>
        public void header(string cosa, XFont font = null, XBrush colore_ = null, int pos = 0, int add_x = 0)
        {
            List<string> split_cosa = cosa.Split(' ').ToList();
            string riga = "";
            XFont font_ = this.font;
            if (font != null)
            {
                font_ = font;
            }
            XBrush colore = XBrushes.Black;
            if (colore_ != null)
            {
                colore = colore_;
            }

            XSize dim = this.gfx.MeasureString(riga, font_, XStringFormats.TopLeft);

            double pos_y = 30;
            XUnit pos_x = XUnit.FromPoint(0);
            switch (pos)
            {
                case 0:
                    pos_x = XUnit.FromPoint(50);
                    break;

                case 1:
                    //pos_x = (this.pdfPage.Width / 2) - (((this.pdfPage.Width - XUnit.FromPoint(100)) / 3) / 2);
                    pos_x = (this.pdfPage.Width / 2) - (XUnit.FromPoint(this.gfx.MeasureString(cosa, font_, XStringFormats.TopLeft).Width) / 2);
                    break;

                case 2:
                    pos_x = (this.pdfPage.Width - XUnit.FromPoint(50)) - ((this.pdfPage.Width - XUnit.FromPoint(100)) / 3);
                    break;
            }

            bool scrivi = false;
            while (true)
            {
                if (XUnit.FromPoint(dim.Width) < (this.pdfPage.Width - XUnit.FromPoint(100)) / 3 & !scrivi)
                {
                    riga += split_cosa[0] + " ";
                    split_cosa.RemoveAt(0);

                    dim = this.gfx.MeasureString(riga, font_, XStringFormats.TopLeft);
                    if (riga.Contains("\n") | split_cosa.Count == 0 |
                                    XUnit.FromPoint(this.gfx.MeasureString(riga + (split_cosa.Count > 0 ? split_cosa[0] : ""),
                                    font_, XStringFormats.TopLeft).Width) > (this.pdfPage.Width - XUnit.FromPoint(100)) / 3)
                    {
                        scrivi = true;
                    }
                }
                else
                {
                    this.gfx.DrawString(riga.Trim(), font_, colore, new XRect((pos_x + XUnit.FromPoint(add_x)).Point, pos_y, ((this.pdfPage.Width - XUnit.FromPoint(100)) / 3).Point, this.pdfPage.Height.Point), XStringFormats.TopLeft);

                    if (split_cosa.Count > 0)
                    {
                        riga = "";
                        dim = this.gfx.MeasureString(riga, font_, XStringFormats.TopLeft);

                        scrivi = false;
                        pos_y += dim.Height + 1;
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            this.pos_y += XUnit.FromPoint(20);
        }
        /// <summary>
        /// Inserice una stringa come footer del PDF
        /// </summary>
        /// <param name="cosa">Stringa da scrivere, per il ritorno a capo inserire l'escape '\n' seguito da uno spazio</param>
        /// <param name="pos">Se 0 posiziona a SX, se 1 posiziona al Centro, se 2 posiziona DX</param>
        public void footer(string cosa, int pos = 0, int add_x = 0)
        {
            List<string> split_cosa = cosa.Split(' ').ToList();
            string riga = "";

            XSize dim = this.gfx.MeasureString(riga, this.font, XStringFormats.TopLeft);

            XUnit pos_x = XUnit.FromPoint(0);
            switch (pos)
            {
                case 0:
                    pos_x = XUnit.FromPoint(50);
                    break;

                case 1:
                    pos_x = (this.pdfPage.Width / 2) - (((this.pdfPage.Width - XUnit.FromPoint(100)) / 3) / 2);
                    break;

                case 2:
                    pos_x = (this.pdfPage.Width - XUnit.FromPoint(50)) - ((this.pdfPage.Width - XUnit.FromPoint(100)) / 3);
                    break;
            }

            bool scrivi = false;
            while (true)
            {
                if (XUnit.FromPoint(dim.Width) < (this.pdfPage.Width - XUnit.FromPoint(100)) / 3 & !scrivi)
                {
                    riga += split_cosa[0] + " ";
                    split_cosa.RemoveAt(0);

                    dim = this.gfx.MeasureString(riga, this.font, XStringFormats.TopLeft);
                    if (riga.Contains("\n") | split_cosa.Count == 0 |
                                    XUnit.FromPoint(this.gfx.MeasureString(riga + (split_cosa.Count > 0 ? split_cosa[0] : ""),
                                    this.font, XStringFormats.TopLeft).Width) > (this.pdfPage.Width - XUnit.FromPoint(100)) / 3)
                    {
                        scrivi = true;
                    }
                }
                else
                {
                    this.gfx.DrawString(riga.Trim(), this.font, XBrushes.Black, new XRect((pos_x + XUnit.FromPoint(add_x)).Point, this.pos_y.Point + 15, ((this.pdfPage.Width - XUnit.FromPoint(100)) / 3).Point, this.pdfPage.Height.Point), XStringFormats.TopLeft);

                    if (split_cosa.Count > 0)
                    {
                        riga = "";
                        dim = this.gfx.MeasureString(riga, this.font, XStringFormats.TopLeft);

                        scrivi = false;
                        this.pos_y += XUnit.FromPoint(dim.Height + 1);
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }
        public void linea(int pos = 0)
        {
            XPen linea = new XPen(XColors.Black, .25);

            XUnit pos_x = XUnit.FromPoint(0);
            switch (pos)
            {
                case 0:
                    pos_x = XUnit.FromPoint(50);
                    break;

                case 1:
                    pos_x = (this.pdfPage.Width / 2) - (((this.pdfPage.Width - (this.pos_x * 2)) / 3) / 2);
                    break;

                case 2:
                    pos_x = (this.pdfPage.Width - XUnit.FromPoint(50)) - ((this.pdfPage.Width - (this.pos_x * 2)) / 3);
                    break;
            }
            this.gfx.DrawLine(linea, pos_x.Point, this.pos_y.Point, (pos_x + ((this.pdfPage.Width - (this.pos_x * 2)) / 3)).Point, this.pos_y.Point);
            this.pos_y += spacing;
        }
        public void riquadro(string cosa, XColor colore_tabella, XBrush colore_font, double larghezza = 0, double altezza = 0)
        {
            XPen pen = new XPen(colore_tabella, .8);
            if (larghezza >= 0 - Costanti.epsilon || larghezza <= 0 + Costanti.epsilon)
            {
                larghezza = (this.pdfPage.Width - (this.pos_x - XUnit.FromPoint(10)) * 2).Point;
            }
            List<string> split_cosa = cosa.Split(' ').ToList();
            XUnit save_pos_y = this.pos_y;


            while (split_cosa.Count > 0)
            {
                split_cosa = this.scrivi_riga(split_cosa, XBrushes.Black);

                if (this.pos_y >= this.pdfPage.Height - (this.margin_y * 2))
                {
                    altezza = (this.pos_y - save_pos_y).Point;
                    this.gfx.DrawRectangle(pen, (pos_x - XUnit.FromPoint(10)).Point, (save_pos_y - XUnit.FromPoint(10)).Point, larghezza + 10, altezza + 10);
                    this.addPage();
                }
            }
            altezza = (this.pos_y - save_pos_y).Point;
            this.gfx.DrawRectangle(pen, (pos_x - XUnit.FromPoint(10)).Point, (save_pos_y - XUnit.FromPoint(10)).Point, larghezza + 10, altezza + 10);
        }
        /// <summary>
        /// Inserisce tabella vuota nel PDF
        /// </summary>
        /// <param name="righe">Numero righe tabella</param>
        /// <param name="colonne">Numero colonne tabella</param>
        /// <param name="colore">Colore bordo tabella</param>
        /// <param name="larghezza">Larghezza colonne tabella, se vuota riempie la pagina</param>
        /// <param name="altezza">Altezza righe tabella, se vuota usa altezza font</param>
        public void tabella(int righe, int colonne, XColor colore, double larghezza = 0, double altezza = 0)
        {
            XPen pen = new XPen(colore);
            if (larghezza >= 0 - Costanti.epsilon || larghezza <= 0 + Costanti.epsilon)
            {
                larghezza = ((this.pdfPage.Width - XUnit.FromPoint(50)) / colonne).Point;
            }
            if (altezza >= 0 - Costanti.epsilon || altezza <= 0 + Costanti.epsilon)
            {
                altezza = this.font.Height + 10;
            }

            for (int riga = 0; riga < righe; riga++)
            {
                double pos_x = 25;
                for (int colonna = 0; colonna < colonne; colonna++)
                {
                    this.gfx.DrawRectangle(pen, pos_x, this.pos_y.Point, larghezza, altezza);
                    pos_x += larghezza;
                }
                this.pos_y += XUnit.FromPoint(altezza);
            }
            this.pos_y += XUnit.FromPoint(altezza);
        }
        /// <summary>
        /// Inserisce e riempie tabella nel PDF
        /// </summary>
        /// <param name="riempi">Lista bidimensionale per rimepire tabella</param>
        /// <param name="colore_tabella">Colore bordo tabella</param>
        /// <param name="colore_font">Colore font</param>
        /// <param name="larghezza">Larghezza colonne tabella, se vuota riempie la pagina</param>
        /// <param name="altezza">Altezza righe tabella, se vuota usa altezza font</param>
        public void tabella(List<List<string>> riempi, XColor colore_tabella, XBrush colore_font, double larghezza = 0, double altezza = 0)
        {
            XPen pen = new XPen(colore_tabella);
            int righe = riempi.Count;
            int colonne = riempi[0].Count;

            if (larghezza >= 0 - Costanti.epsilon || larghezza <= 0 + Costanti.epsilon)
            {
                larghezza = ((this.pdfPage.Width - XUnit.FromPoint(100)) / colonne).Point;
            }
            if (altezza >= 0 - Costanti.epsilon || altezza <= 0 + Costanti.epsilon)
            {
                altezza = this.font.Height + 10;
            }

            for (int riga = 0; riga < righe; riga++)
            {
                double pos_x = 50;
                int quante_righe = 0;

                List<List<string>> riga_out = new List<List<string>>();
                List<string> nuova_riga = new List<string>();

                for (int colonna = 0; colonna < colonne; colonna++)
                {
                    string stringa_misura = string.Empty;
                    List<string> righe_divise = riempi[riga][colonna].Split('\n').ToList();

                    //foreach (string parole_riga in righe_divise)
                    for (int indice_parola = 0; indice_parola < righe_divise.Count; indice_parola++)
                    {
                        string parole_riga = righe_divise[indice_parola];
                        List<string> p = parole_riga.Split(' ').ToList();

                        while (true)
                        {
                            p.ForEach(x => stringa_misura += x + " ");

                            XSize dim = this.gfx.MeasureString(stringa_misura, font, XStringFormats.Center);
                            if (dim.Width >= larghezza - 10)
                            {
                                nuova_riga.Insert(0, p[p.Count - 1]);
                                p.RemoveAt(p.Count - 1);

                                stringa_misura = string.Empty;
                                continue;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (riga_out.Count - 1 < colonna)
                        {
                            riga_out.Add(new List<string>());
                        }
                        riga_out[colonna].Add(stringa_misura);

                        if (quante_righe < riga_out[colonna].Count)
                        {
                            quante_righe = riga_out[colonna].Count;
                        }

                        if (nuova_riga.Count > 0)
                        {
                            righe_divise[indice_parola] = righe_divise[indice_parola].Substring(stringa_misura.Length).Trim();
                            nuova_riga = new List<string>();
                            indice_parola--;
                        }

                        stringa_misura = string.Empty;
                    }
                }
                altezza = (this.font.Height + int.Parse(this.spacing.Point.ToString())) * quante_righe;

                entra(2);
                foreach (List<string> s in riga_out)
                {
                    this.gfx.DrawRectangle(pen, pos_x, this.pos_y.Point, larghezza, altezza + 10);

                    XUnit tmp_y = this.pos_y + (quante_righe > 1 ? ((XUnit.FromPoint(altezza / 2) - XUnit.FromPoint(15)) * -1) : XUnit.FromPoint(5));
                    foreach (string a in s)
                    {
                        this.gfx.DrawString(a, this.font, colore_font, new XRect(pos_x, tmp_y.Point, larghezza, altezza), XStringFormats.Center);

                        tmp_y += XUnit.FromPoint(this.font.Height) + this.spacing;
                    }
                    pos_x += larghezza;
                }

                this.pos_y += XUnit.FromPoint(altezza + 10);
            }
            this.pos_y += XUnit.FromPoint(20);
        }
        public void salva(string path, string nome_file = "")
        {
            if (nome_file.Length > 0)
            {
                path = path.EndsWith("\\") ? path : path + "\\";
                nome_file = nome_file.EndsWith(".pdf") ? nome_file : nome_file + ".pdf";
            }

            this.pdfDocument.Save(path + nome_file);

            this.pdfDocument = new PdfDocument();
            this.pdfPage = pdfDocument.AddPage();

            this.gfx = XGraphics.FromPdfPage(pdfPage);
            this.tf = new XTextFormatter(gfx);
            this.pos_y = XUnit.FromPoint(50);
            XFont font = new XFont("Tahoma", 12);
            this.pagina = 1;
            this.header("Pag. " + pagina, font, null, 2, 150);
        }
        public void addPage()
        {
            this.pagina += 1;
            this.pos_y = this.pdfPage.Height - XUnit.FromPoint(40);
            this.footer("Segue >>", 2, 150);

            this.pdfPage = this.pdfDocument.AddPage();
            this.gfx = XGraphics.FromPdfPage(pdfPage);
            this.tf = new XTextFormatter(gfx);
            this.tf.Font = this.font;
            XFont font = new XFont("Tahoma", 12);

            this.pos_y = XUnit.FromPoint(50);
            this.header("Pag. " + pagina, font, null, 2, 150);
        }
        public bool entra(int numero_righe)
        {
            bool entra = false;
            XUnit altezza_riga = XUnit.FromPoint(this.font.Height) + spacing;

            if (this.pos_y + (altezza_riga * numero_righe) > this.pdfPage.Height - XUnit.FromPoint(70))
            {
                this.addPage();
            }
            else
            {
                entra = true;
            }

            return entra;
        }
    }
    public static class OCF
    {
        public static ChromeDriver driver = null;
        static WebDriverWait wait;

        public static Dictionary<string, string> cerca(string cognome, string nome, string data_nascita)
        {
            List<string> result = new List<string>();
            Dictionary<string, string> new_result = new Dictionary<string, string>();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            new Actions(driver).SendKeys(Keys.PageDown).Perform();
            Thread.Sleep(2000);
            try
            {
                wait.Timeout = TimeSpan.FromSeconds(3);
                wait.Until(d => d.FindElement(By.XPath("//input[@name='impliedsubmit']"))).Click();  // CHIUSURA POPUP COOKIE
                wait.Timeout = TimeSpan.FromSeconds(10);
                Thread.Sleep(500);
            }
            catch (Exception)
            {
            }

            wait.Until(d => d.FindElement(By.Id("clearButton"))).Click();
            Thread.Sleep(1000);

            wait.Until(d => d.FindElement(By.Id("cognome"))).SendKeys(cognome);
            Thread.Sleep(500);
            wait.Until(d => d.FindElement(By.Id("nome"))).SendKeys(nome);
            Thread.Sleep(500);
            string captcha = string.Empty;
            while (true)
            {
                captcha = new InputBox("Captcha", "Captcha OCF").result;

                wait.Until(d => d.FindElement(By.Id("captchaInput"))).SendKeys(captcha);
                Thread.Sleep(1000);

                new Actions(driver).SendKeys(Keys.PageDown).Perform();
                Thread.Sleep(500);

                wait.Until(d => d.FindElement(By.Id("submitRicercaConsulenteButton"))).Click();
                Thread.Sleep(3000);
                new Actions(driver).SendKeys(Keys.End).Perform();
                Thread.Sleep(500);

                try
                {
                    if (driver.FindElement(By.ClassName("lfr-alert-container")).Displayed)
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    break;
                }
            }

            IList<IWebElement> risultato_ricerca = wait.Until(d => d.FindElements(By.XPath("//button[contains(@class, 'showDetail_dettaglio_domanda')]")));
            foreach (IWebElement elemento in risultato_ricerca)
            {
                while (true)
                {
                    try
                    {
                        elemento.Click();
                        Thread.Sleep(1000);
                        break;
                    }
                    catch (Exception ex)
                    {
                        new Actions(driver).SendKeys(Keys.PageUp).Perform();
                        Thread.Sleep(1000);
                        continue;
                    }
                }

                IList<IWebElement> dettaglio = new List<IWebElement>();
                while (true)
                {
                    dettaglio = wait.Until(d => d.FindElements(By.XPath("//div[contains(@class, 'customize-cell')]")));
                    // dettaglio[0] -> NOMINATIVO
                    // dettaglio[1] -> SEZIONE
                    // dettaglio[2] -> DATA NASCITA
                    // dettaglio[3] -> LUOGO NASCITA
                    // dettaglio[4] -> DOMICILIO
                    // dettaglio[5] -> INDIRIZZO
                    // dettaglio[6] -> STATO

                    if (dettaglio.Count == 0)
                    {
                        Thread.Sleep(1000);
                        continue;
                    }
                    else if (dettaglio.Count >= 7)
                    {
                        break;
                    }
                    else
                    {
                        Thread.Sleep(1000);
                        continue;
                    }
                }

                foreach (IWebElement elem in dettaglio)
                {
                    string descr = Utils.sina(elem.Text, ":").Trim().ToUpper();
                    string valor = Utils.desa(elem.Text, ":").Trim().ToUpper();

                    new_result.Add(descr, valor);
                    if (elem.Text == "")
                    {
                        result.Add("N.D.");
                    }
                    else
                    {
                        result.Add(elem.Text.ToUpper().Trim());
                    }
                }

                if (new_result["DATA DI NASCITA"] == data_nascita)
                {
                    return new_result;
                }
                else
                {
                    new_result = new Dictionary<string, string>();
                }
            }
            //return new List<string>();
            return new Dictionary<string, string>();
        }
        public static List<Dictionary<string, string>> storico()
        {
            List<List<string>> result = new List<List<string>>();
            List<Dictionary<string, string>> new_result = new List<Dictionary<string, string>>();

            List<string> tmp = new List<string>();
            Dictionary<string, string> new_tmp = new Dictionary<string, string>();
            List<List<string>> storico = new List<List<string>>();

            while (true)
            {
                result = new List<List<string>>();
                tmp = new List<string>();

                storico = HTML.CaptureTable("[@id='tableContainerStorico']");
                if (storico.Count > 0)
                {
                    if (storico[0].Count == 6)
                    {
                        Thread.Sleep(2000);
                        storico = HTML.CaptureTable("[@id='tableContainerStorico']");
                        break;
                    }
                }
                else
                {
                    Thread.Sleep(500);
                    continue;
                }
            }

            foreach (List<string> x in storico)
            {
                string dett = string.Empty;
                int i = 0;
                foreach (string s in x)
                {
                    dett = s;
                    string descr = string.Empty;

                    switch (i)
                    {
                        case 0:
                            descr = "STATO";
                            break;

                        case 1:
                            descr = "SEZIONE ALBO";
                            break;

                        case 2:
                            descr = "DELIBERA";
                            break;

                        case 3:
                            descr = "DATA DELIBERA";
                            break;

                        case 4:
                            descr = "DATA EFFICACIA";
                            break;

                        case 5:
                            descr = "ENTE";
                            break;
                    }

                    if (dett != "")
                    {
                        if (dett.Contains("<span"))
                        {
                            dett = Utils.ExtractBetween(dett, ">", "</span");
                        }
                        if (dett.Contains("<b>"))
                        {
                            dett = Utils.ExtractBetween(dett, "b>", "</b");
                        }
                        tmp.Add(dett.ToUpper());
                    }
                    else
                    {
                        tmp.Add("N.D.");
                    }
                    new_tmp.Add(descr, dett);

                    i += 1;
                }

                if (tmp.Count == 1)
                {
                    Thread.Sleep(3000);
                    break;
                }

                result.Add(tmp);
                new_result.Add(new_tmp);

                tmp = new List<string>();
                new_tmp = new Dictionary<string, string>();
            }

            return new_result;
        }
        public static List<Dictionary<string, string>> intermediari()
        {
            int tentativi = 0;
            List<List<string>> result = new List<List<string>>();
            List<Dictionary<String, string>> new_result = new List<Dictionary<string, string>>();

            List<string> tmp = new List<string>();
            Dictionary<string, string> new_tmp = new Dictionary<string, string>();

            List<List<string>> intermediari = new List<List<string>>();
            while (true)
            {
                IWebElement div_tab = driver.FindElement(By.Id("divContainer"));
                intermediari = HTML.CaptureTable("", div_tab);

                if (intermediari.Count > 0)
                {
                    if (intermediari[0].Count >= 3)
                    {
                        break;
                    }
                }
                else
                {
                    if (tentativi < 10)
                    {
                        tentativi += 1;
                        Thread.Sleep(1000);
                        continue;
                    }
                    break;
                }
            }

            foreach (List<string> x in intermediari)
            {
                string dett = string.Empty;
                int i = 0;

                foreach (string s in x)
                {
                    string descr = string.Empty;
                    switch (i)
                    {
                        case 0:
                            descr = "SOGGETTO";
                            break;

                        case 1:
                            descr = "DATA INIZIO";
                            break;

                        case 2:
                            descr = "DATA FINE";
                            break;

                        default:
                            continue;
                    }

                    dett = s;
                    if (dett != "")
                    {
                        if (dett.Contains("<span"))
                        {
                            dett = Utils.ExtractBetween(dett, ">", "</span");
                        }
                        if (dett.Contains("<b>"))
                        {
                            dett = Utils.ExtractBetween(dett, "b>", "</b");
                        }
                        tmp.Add(dett.ToUpper());
                    }
                    else
                    {
                        tmp.Add("N.D.");
                    }
                    new_tmp.Add(descr, dett);

                    i += 1;
                }
                result.Add(tmp);
                new_result.Add(new_tmp);

                tmp = new List<string>();
                new_tmp = new Dictionary<string, string>();
            }
            return new_result;
        }
        public static void collega_driver(ChromeDriver browser)
        {
            driver = browser;

            HTML.browser = "CR";
            HTML.chrome = browser;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
    }
    public static class OAM
    {
        public static ChromeDriver driver = null;
        static WebDriverWait wait;

        /// <summary>
        /// <br>Cerca in OAM il cofice fiscale indicato</br> 
        /// <br>il metodo apre la finestra dettaglio del soggetto</br>
        /// </summary>
        /// <param name="cf"></param>
        /// <returns>true o false se trova oppure no il codice fiscale</returns>
        public static bool ricerca(string cf)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            try
            {
                wait.Until(d => d.FindElement(By.LinkText("Ricerca"))).Click();
            }
            catch
            {
            }

            Thread.Sleep(1000);
            while (true)
            {
                try
                {
                    wait.Until(d => d.FindElement(By.Id("resetta-filtri-elenchi-agenti"))).Click();
                    break;
                }
                catch
                {
                    new Actions(driver).SendKeys(Keys.End).Perform();
                }
            }
            Thread.Sleep(500);

            int quanti = 0;
            while (true)
            {
                quanti = 0;
                wait.Until(d => d.FindElement(By.Name("CODICE_FISCALE"))).SendKeys(cf);
                Thread.Sleep(500);
                while (true)
                {
                    try
                    {
                        wait.Until(d => d.FindElement(By.Id("filtri-elenchi-agenti"))).Click();
                        break;
                    }
                    catch
                    {
                        new Actions(driver).SendKeys(Keys.PageDown).Perform();
                    }
                }
                Thread.Sleep(4000);

                while (true)
                {
                    try
                    {
                        quanti = int.Parse(wait.Until(d => d.FindElement(By.XPath("//span[@class='totale-risultati-elenchi']"))).GetAttribute("innerHTML"));
                        break;
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                if (quanti == 1)
                {
                    wait.Until(d => d.FindElement(By.Id("dettaglio_iscritto"))).Click();
                    Thread.Sleep(1000);
                    return true;
                }
                else
                {
                    if (quanti == 0)
                    {
                        break;
                    }

                    wait.Until(d => d.FindElement(By.LinkText("Ricerca"))).Click();
                    Thread.Sleep(2000);
                    while (true)
                    {
                        try
                        {
                            wait.Until(d => d.FindElement(By.Id("resetta-filtri-elenchi-agenti"))).Click();
                            break;
                        }
                        catch
                        {
                            new Actions(driver).SendKeys(Keys.End).Perform();
                        }
                    }
                }
            }
            return false;
        }
        public static Dictionary<string, string> ricerca_collab(string cf)
        {
            //List<string> ret = new List<string>();
            Dictionary<string, string> ret = new Dictionary<string, string>();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            try
            {
                wait.Until(d => d.FindElement(By.LinkText("Ricerca"))).Click();
            }
            catch
            {
            }

            while (true)
            {
                try
                {
                    Thread.Sleep(1000);
                    wait.Until(d => d.FindElement(By.Id("cancella-collaboratore"))).Click();
                    Thread.Sleep(500);

                    break;
                }
                catch
                {
                    new Actions(driver).SendKeys(Keys.End).Perform();
                    continue;
                }
            }

            int quanti = 0;
            while (true)
            {
                quanti = 0;
                wait.Until(d => d.FindElement(By.Name("COLLABORATORE_CODICE_FISCALE"))).SendKeys(cf);
                Thread.Sleep(500);
                wait.Until(d => d.FindElement(By.Id("ricerca-collaboratore"))).Click();
                Thread.Sleep(4000);

                quanti = int.Parse(wait.Until(d => d.FindElement(By.XPath("//span[@class='totale-risultati-collaboratori']"))).GetAttribute("innerHTML"));
                if (quanti == 1)
                {
                    //string tmp = wait.Until(d => d.FindElement(By.Id("dettaglio_collaboratore"))).GetAttribute("innerHTML");
                    string tmp = wait.Until(d => d.FindElement(By.Id("dettaglio_collaboratore"))).Text;

                    ret.Add("NOMINATIVO", tmp.Substring(0, tmp.IndexOf(", CF")).Trim());
                    tmp = tmp.Substring(tmp.IndexOf(", CF") + 4);
                    ret.Add("CODICE FISCALE", tmp.Substring(0, tmp.IndexOf(", DIPENDENTE")).Trim());
                    tmp = tmp.Substring(tmp.IndexOf(", DAL") + 5);
                    ret.Add("INIZIO COLLABORAZIONE", tmp.Substring(0, tmp.IndexOf("\n")).Trim());
                    tmp = tmp.Substring(tmp.IndexOf(":") + 1);
                    ret.Add("SOCIETA", tmp.Substring(0, tmp.IndexOf("|")).Trim());
                    tmp = tmp.Substring(tmp.IndexOf(":") + 1);
                    ret.Add("ELENCO", tmp.Substring(0, tmp.IndexOf("| ISCRIZIONE")).Trim());
                    tmp = tmp.Substring(tmp.IndexOf(":") + 1);
                    ret.Add("NUMERO ISCRIZIONE", tmp.Trim());

                    Thread.Sleep(1000);

                    return ret;
                }
                else
                {
                    if (quanti == 0)
                    {
                        break;
                    }
                    wait.Until(d => d.FindElement(By.LinkText("Ricerca"))).Click();
                    Thread.Sleep(2000);
                    wait.Until(d => d.FindElement(By.Id("cancella-collaboratore"))).Click();
                }
            }
            return ret;
        }
        /// <summary>
        /// Cattura dettagli soggetto sulla prima schermata
        /// </summary>
        /// <returns>
        /// <br>[0] TIPO ELENCO</br>
        /// <br>[1] NUMERO ISCRIZIONE</br>
        /// <br>[2] STATO ISCRIZIONE</br>
        /// <br>[3] AUTORIZZATO AD OPERARE</br>
        /// <br>[4] DENOMINAZIONE</br>
        /// <br>[5] NDG | SESSO</br>
        /// <br>[6] CODICE FISCALE</br>
        /// <br>[7] DATA COSTITUZIONE | DATA NASCITA</br>
        /// <br>[8] PEC | COMUNE NASCITA</br>
        /// <br>[9] DENOMINAZIONE RAPPRESENTANTE LEGALE | PROVINCIA NASCITA</br>
        /// <br>[10] SESSO RAPPRESENTANTE LEGALE | CITTADINANZA</br>
        /// <br>[11] CODICE FISCALE RAPPRESENTANTE LEGALE | DITTA</br>
        /// <br>[12] PROVINCIA NASCITA RAPPRESENTANTE LEGALE | PEC</br>
        /// <br>[13] COMUNE NASCITA RAPPRESENTANTE LEGALE</br>
        /// <br>[14] DATA NASCITA RAPPRESENTANTE LEGALE</br>
        /// <br>[15] INIZIO INCARICO</br>
        /// </returns>
        public static Dictionary<string, string> cattura_dettaglio()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            List<string> result = new List<string>();
            Dictionary<string, string> new_result = new Dictionary<string, string>();

            IWebElement elem;

            bool ditta = false;
            string descr = string.Empty;
            string campo = string.Empty;

            string tmp = string.Empty;

            int indice_elem = 0;
            IWebElement dettagli = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'dettaglio_generico')]")));
            foreach (IWebElement li in dettagli.FindElements(By.XPath(".//li")))
            {
                try
                {
                    elem = li.FindElement(By.XPath(".//span"));
                }
                catch (Exception)
                {
                    try
                    {
                        elem = li.FindElement(By.XPath(".//div[@class='denominazione-dati-anagrafici']"));
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                if (descr.Contains("mailto"))
                {
                    descr = Utils.ExtractBetween(descr, ">", "</a>");
                }
                campo = li.Text;


                switch (indice_elem)
                {
                    case 10:
                        if (!campo.Contains("Cittadinanza"))
                        {
                            result.Add("");
                        }
                        break;

                    case 11:
                        if (!campo.Contains("Ditta"))
                        {
                            result.Add("");
                        }
                        break;
                }

                if (campo.Contains("\r\n"))
                {

                    if (new_result.ContainsKey("DENOMINAZIONE") & campo.Substring(0, campo.IndexOf("\r\n")).ToUpper() == "COGNOME E NOME")
                    {
                        tmp = "_RAPP_LEGALE";
                    }
                    if (!new_result.ContainsKey(campo.Substring(0, campo.IndexOf("\r\n")).ToUpper() + tmp))
                    {
                        new_result.Add(campo.Substring(0, campo.IndexOf("\r\n")).ToUpper().Trim() + tmp, campo.Substring(campo.IndexOf("\r\n")).Trim());
                    }
                }

                switch (campo)
                {
                    case string dettaglio when dettaglio.Contains("Tipo Elenco") |
                                               dettaglio.Contains("Numero Iscrizione") |
                                               dettaglio.Contains("Stato") |
                                               dettaglio.Contains("Autorizzato ad operare"):
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("Denominazione"):
                        ditta = true;
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("Natura giuridica") |
                                               dettaglio.Contains("Codice Fiscale") |
                                               dettaglio.Contains("Data Costituzione"):
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("Cognome e Nome") |
                                               dettaglio.Contains("Sesso") |
                                               dettaglio.Contains("Codice Fiscale") |
                                               dettaglio.Contains("Data di nascita") |
                                               dettaglio.Contains("Comune di nascita") |
                                               dettaglio.Contains("Provincia di nascita") |
                                               dettaglio.Contains("Cittadinanza"):
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("Inizio incarico"):
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("Ditta"):
                        ditta = true;
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("PEC"):
                        if (!ditta)
                        {
                            result.Add("");
                        }
                        result.Add(descr);
                        break;

                    case string dettaglio when dettaglio.Contains("Indirizzo") |
                                               dettaglio.Contains("CAP") |
                                               dettaglio.Contains("Cap") |
                                               dettaglio.Contains("Comune") |
                                               dettaglio.Contains("Provincia"):
                        result.Add(descr);
                        break;
                }
            }
            return new_result;
        }
        public static Dictionary<string, string> cattura_sedi()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            List<string> result = new List<string>();
            Dictionary<string, string> new_result = new Dictionary<string, string>();

            string tmp = "_DIR_GENERALE";

            int tentativo = 0;

            while (true)
            {
                new Actions(driver).SendKeys(Keys.End).Perform();
                Thread.Sleep(1000);
                try
                {
                    wait.Until(d => d.FindElement(By.LinkText("Sedi"))).Click();
                    Thread.Sleep(3000);
                    break;
                }
                catch (Exception)
                {
                    if (tentativo > 5)
                    {
                        return new_result;
                    }
                    tentativo += 1;
                    continue;
                }
            }

            while (true)
            {
                IWebElement sedi = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'lista_sedi')]")));
                foreach (IWebElement li in sedi.FindElements(By.XPath(".//li")))
                {
                    try
                    {
                        string descr = li.Text.Trim();
                        switch (descr)
                        {
                            case "Direzione Generale":
                                tmp = "_DIR_GENERALE";
                                break;

                            case "Sede Italiana":
                                tmp = "_ITALIA";
                                break;

                            default:
                                result.Add(descr);

                                new_result.Add(descr.Substring(0, descr.IndexOf("\r\n")).ToUpper() + tmp, descr.Substring(descr.IndexOf("\r\n")).Trim());
                                break;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                if (result.Count == 0)
                {
                    continue;
                }
                else
                {
                    wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                    Thread.Sleep(1000);
                    return new_result;
                }
            }
        }
        public static List<Dictionary<string, string>> cattura_mandati(string diretti_indiretti, bool rapporti_cessati = false)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            Dictionary<string, string> intermed = new Dictionary<string, string>();

            List<List<string>> result = new List<List<string>>();
            List<Dictionary<string, string>> new_result = new List<Dictionary<string, string>>();

            while (true)
            {
                Thread.Sleep(1000);
                while (true)
                {
                    try
                    {
                        wait.Until(d => d.FindElement(By.LinkText("Mandati " + diretti_indiretti))).Click();
                        Thread.Sleep(1000);
                        break;
                    }
                    catch (ElementClickInterceptedException)
                    {
                        new Actions(driver).SendKeys(Keys.End).Perform();
                        continue;
                    }
                    catch (Exception)
                    {
                        return new_result;
                    }
                }

                if (HTML.vedisece("MANDATI " + diretti_indiretti.ToUpper(), driver))
                {
                    break;
                }
                else
                {
                    wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                    Thread.Sleep(1000);
                    new Actions(driver).SendKeys(Keys.End).Perform();
                    continue;
                }
            }

            Thread.Sleep(3000);
            IWebElement mandati = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'mandati_" + diretti_indiretti.ToLower() + "')]")));
            foreach (IWebElement li in mandati.FindElements(By.XPath(".//li")))
            {
                string descr = li.Text;
                string valor = li.Text;
                try
                {
                    if (descr.Contains("\n"))
                    {
                        descr = descr.Substring(0, descr.IndexOf("\n")).Trim().ToUpper();
                        valor = valor.Substring(valor.IndexOf("\n")).Trim();
                    }

                    IWebElement elem = li.FindElement(By.XPath(".//span"));

                    if (descr.Contains("Prodotti e Attività"))
                    {
                        new_result.Add(intermed);
                        intermed = new Dictionary<string, string>();
                        continue;
                    }

                    intermed.Add(descr, valor);
                }
                catch (ArgumentException)
                {
                    if (descr == "DENOMINAZIONE" | descr == "CODICE FISCALE")
                    {
                        intermed.Add(descr + "_INTERMED", valor);
                    }
                }
                catch (Exception)
                {
                    try
                    {
                        IWebElement elem = li.FindElement(By.XPath(".//div[@class='denominazione-dati-anagrafici']"));
                        descr = elem.GetAttribute("innerHTML");
                        descr = descr.Replace("&nbsp;", "");
                        descr = descr.Replace("&amp;", "");
                        descr = descr.Replace("<em>", "");
                        descr = descr.Replace("</em>", "");

                        //intermed.Add(descr);
                    }
                    catch (Exception)
                    {
                        string eccolo = li.GetAttribute("innerHTML");
                        if (li.GetAttribute("innerHTML").Contains("Rapporti Cessati"))
                        {
                            if (rapporti_cessati)
                            {
                                new_result = new List<Dictionary<string, string>>();
                                continue;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
            Thread.Sleep(1000);
            return new_result;
        }
        public static List<List<string>> cattura_mandati_diretti()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            List<string> intermed = new List<string>();
            List<List<string>> result = new List<List<string>>();
            List<Dictionary<string, string>> new_result = new List<Dictionary<string, string>>();

            new Actions(driver).SendKeys(Keys.End).Perform();
            Thread.Sleep(200);
            try
            {
                wait.Until(d => d.FindElement(By.LinkText("MANDATI DIRETTI"))).Click();
                Thread.Sleep(1000);
            }
            catch (Exception)
            {
                return result;
            }

            IWebElement mandati = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'mandati_diretti')]")));
            foreach (IWebElement li in mandati.FindElements(By.XPath(".//li")))
            {
                try
                {
                    IWebElement elem = li.FindElement(By.XPath(".//span"));
                    string descr = elem.GetAttribute("innerHTML");
                    descr = descr.Replace("&nbsp;", "");
                    descr = descr.Replace("<em>", "");
                    descr = descr.Replace("</em>", "");

                    intermed.Add(descr);
                }
                catch (Exception)
                {
                    try
                    {
                        IWebElement elem = li.FindElement(By.XPath(".//div[@class='denominazione-dati-anagrafici']"));
                        string descr = elem.GetAttribute("innerHTML");
                        descr = descr.Replace("&nbsp;", "");
                        descr = descr.Replace("<em>", "");
                        descr = descr.Replace("</em>", "");

                        intermed.Add(descr);
                    }
                    catch (Exception)
                    {
                        if (li.Text == "Prodotti e Attività")
                        {
                            result.Add(intermed);
                            intermed = new List<string>();
                        }
                        else if (li.Text == "Rapporti Cessati")
                        {
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
            Thread.Sleep(1000);
            return result;
        }
        public static List<List<string>> cattura_mandati_indiretti()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            List<string> intermed = new List<string>();
            List<List<string>> result = new List<List<string>>();

            new Actions(driver).SendKeys(Keys.End).Perform();
            Thread.Sleep(200);
            try
            {
                wait.Until(d => d.FindElement(By.LinkText("MANDATI INDIRETTI"))).Click();
                Thread.Sleep(1000);
            }
            catch (Exception)
            {
                return result;
            }

            IWebElement mandati = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'mandati_indiretti')]")));
            foreach (IWebElement li in mandati.FindElements(By.XPath(".//li")))
            {
                try
                {
                    IWebElement elem = li.FindElement(By.XPath(".//span"));
                    string descr = elem.GetAttribute("innerHTML");
                    descr = descr.Replace("&nbsp;", "");
                    descr = descr.Replace("<em>", "");
                    descr = descr.Replace("</em>", "");

                    intermed.Add(descr);
                }
                catch (Exception)
                {
                    try
                    {
                        IWebElement elem = li.FindElement(By.XPath(".//div[@class='denominazione-dati-anagrafici']"));
                        string descr = elem.GetDomProperty("innerHTML");
                        descr = descr.Replace("&nbsp;", "");
                        descr = descr.Replace("<em>", "");
                        descr = descr.Replace("</em>", "");

                        intermed.Add(descr);
                    }
                    catch (Exception)
                    {
                        if (li.Text == "Prodotti e Attività")
                        {
                            result.Add(intermed);
                            intermed = new List<string>();
                        }
                        else if (li.Text == "Rapporti Cessati")
                        {
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
            Thread.Sleep(1000);
            return result;
        }
        public static List<List<string>> cattura_prodotti(int indice_intermed, string mandati = "Diretti")
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            List<string> tmp = new List<string>();
            List<List<string>> result = new List<List<string>>();

            new Actions(driver).SendKeys(Keys.End).Perform();
            Thread.Sleep(200);
            while (true)
            {
                try
                {
                    wait.Until(d => d.FindElement(By.LinkText("Mandati " + mandati))).Click();
                    break;
                }
                catch (ElementClickInterceptedException)
                {
                    new Actions(driver).SendKeys(Keys.PageDown).Perform();
                    Thread.Sleep(1000);
                    continue;
                }
                catch (WebDriverTimeoutException)
                {
                    Thread.Sleep(1000);
                    continue;
                }
            }
            Thread.Sleep(1500);

            if (mandati.ToUpper() == "DIRETTI")
            {
                while (true)
                {
                    try
                    {
                        Thread.Sleep(1000);
                        IList<IWebElement> l = wait.Until(d => d.FindElements(By.Id("prodotti_mandato_diretto")));
                        if (l.Count > 0)
                        {
                            l[indice_intermed].Click();
                            break;
                        }
                    }
                    catch (ElementClickInterceptedException)
                    {
                        new Actions(driver).SendKeys(Keys.PageDown).Perform();
                        Thread.Sleep(1000);
                        continue;
                    }
                }
            }
            else
            {
                while (true)
                {
                    try
                    {
                        wait.Until(d => d.FindElements(By.Id("prodotti_mandato")))[indice_intermed].Click();
                        break;
                    }
                    catch (ElementClickInterceptedException)
                    {
                        new Actions(driver).SendKeys(Keys.PageDown).Perform();
                        Thread.Sleep(1000);
                        continue;
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        continue;
                    }
                }
            }
            Thread.Sleep(2000);

            IWebElement lista = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'elenco_prodotti')]")));
            foreach (IWebElement li in lista.FindElements(By.XPath(".//li")))
            {
                string prod = li.GetDomProperty("innerHTML");
                prod = prod.Substring(prod.IndexOf("CODICE:"));
                tmp.Add(Utils.ExtractBetween(prod, "CODICE:", "</p>"));
                prod = prod.Substring(prod.IndexOf("DESCRIZIONE:"));
                tmp.Add(Utils.ExtractBetween(prod, "DESCRIZIONE:", "</p>"));
                result.Add(tmp);
                tmp = new List<string>();
            }
            wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
            Thread.Sleep(1000);
            wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
            Thread.Sleep(1000);
            return result;
        }
        /// <summary>
        /// Cattura dipendenti da OAM
        /// </summary>
        /// <returns>
        ///     <br>[0] - DENOM</br>
        ///     <br>[1] - LUOGO NASCITA</br>
        ///     <br>[2] - DATA NASCITA</br>
        ///     <br>[3] - CODICE FISCALE</br>
        ///     <br>[4] - SESSO</br>
        ///     <br>[5] - INIZIO COLLABORAZIONE</br>
        ///     <br>[6] - NUMERO ISCRIZIONE</br>
        ///     <br>[7] - AMMINISTRATORE</br>    
        ///     <br>[8] - STATO ISCRIZIONE</br>    
        /// </returns>
        public static List<Dictionary<string, string>> cattura_dipendenti()
        {
            while (true)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                List<string> dipen = new List<string>();
                List<List<string>> result = new List<List<string>>();
                List<Dictionary<string, string>> new_result = new List<Dictionary<string, string>>();

                while (true)
                {
                    Thread.Sleep(3000);
                    while (true)
                    {
                        try
                        {
                            wait.Until(d => d.FindElement(By.LinkText("Dipendenti Collaboratori"))).Click();
                            Thread.Sleep(4000);
                            break;
                        }
                        catch (ElementClickInterceptedException)
                        {
                            new Actions(driver).SendKeys(Keys.End).Perform();
                            continue;
                        }
                        catch (Exception)
                        {
                            return new_result;
                        }
                    }

                    if (HTML.vedisece("DIPENDENTI E COLLABORATORI", driver))
                    {
                        break;
                    }
                    else
                    {
                        wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                        Thread.Sleep(1000);
                        new Actions(driver).SendKeys(Keys.End).Perform();
                        continue;
                    }
                }

                string dip = wait.Until(d => d.FindElement(By.XPath("//span[@class='total_dipendenti_attivi']"))).GetDomProperty("innerHTML");
                if (dip.Trim() == "")
                {
                    wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                    Thread.Sleep(1000);

                    return new_result;
                }

                wait.Until(d => d.FindElement(By.XPath("//span[@class='total_dipendenti_attivi']"))).Click();
                Thread.Sleep(3000);
                int quanti_dipendenti = int.Parse(dip.Replace("(", "").Replace(")", ""));

                if (quanti_dipendenti > 0)
                {
                    int indice_dipendente = 0;
                    //IWebElement lista_dipen = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'dipendenti_collaboratori_attivi')]")));
                    while (true)
                    {
                        IList<IWebElement> lista_dipen = wait.Until(d => d.FindElements(By.XPath("//ul[contains(@class, 'dipendenti_collaboratori_attivi')]/li")));
                        if (indice_dipendente >= lista_dipen.Count)
                        {
                            break;
                        }
                        IWebElement li = lista_dipen[indice_dipendente];

                        Dictionary<string, string> dict = new Dictionary<string, string>();
                        Dictionary<string, string> dettagli_dipen = new Dictionary<string, string>();

                        List<string> tmp_ = li.FindElement(By.Id("dettaglio_iscritto")).Text.Split('\n').ToList();
                        string tmp = "";

                        foreach (string s in tmp_)
                        {
                            string descr = string.Empty;
                            try
                            {
                                descr = s.Substring(0, s.IndexOf(":"));
                            }
                            catch { }
                            string valor = s.Substring(s.IndexOf(":") + 1);

                            if (descr == "")
                            {
                                dict.Add("NOMINATIVO", valor.Trim());
                            }
                            else
                            {
                                string x = string.Empty;
                                switch (descr)
                                {
                                    case "LUOGO DI NASCITA":
                                        valor = valor.Substring(0, valor.IndexOf("-")).Trim();
                                        dict.Add(descr, valor);

                                        x = s.Substring(s.IndexOf("-") + 1);
                                        descr = x.Substring(0, x.IndexOf(":")).Trim();
                                        valor = x.Substring(x.IndexOf(":") + 1).Trim();
                                        dict.Add(descr, valor);

                                        break;

                                    case "CODICE FISCALE":
                                        valor = valor.Substring(0, valor.IndexOf(" SESSO")).Trim();
                                        dict.Add(descr, valor);

                                        x = s.Substring(s.IndexOf(" SESSO"));
                                        descr = x.Substring(0, x.IndexOf(":")).Trim();
                                        valor = x.Substring(x.IndexOf(":") + 1).Trim();
                                        dict.Add(descr, valor);

                                        break;

                                    default:
                                        try
                                        {
                                            dict.Add(descr, valor.Trim());
                                        }
                                        catch { }

                                        break;
                                }
                            }
                        }

                        if (dict.ContainsKey("NUMERO ISCRIZIONE"))
                        {
                            while (true)
                            {
                                try // CLICCO SU DIPENDENTE
                                {
                                    li.Click();
                                    break;
                                }
                                catch (ElementClickInterceptedException)
                                {
                                    new Actions(driver).SendKeys(Keys.PageDown).Perform();
                                    Thread.Sleep(1000);
                                    continue;
                                }
                            }
                            Thread.Sleep(1500);
                            dettagli_dipen = cattura_dettaglio();
                            wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                            Thread.Sleep(3000);
                            int tentativi = 0;
                            while (true)
                            {
                                try
                                {
                                    wait.Until(d => d.FindElement(By.XPath("//span[@class='total_dipendenti_attivi']"))).Click();
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    Thread.Sleep(500);
                                    tentativi += 1;
                                    if (tentativi > 10)
                                    {
                                        return null;
                                    }
                                    continue;
                                }
                            }
                            if (dettagli_dipen.ContainsKey("STATO"))
                            {
                                dict.Add("STATO", dettagli_dipen["STATO"]);
                            }
                        }

                        new_result.Add(dict);
                        indice_dipendente += 1;
                    }

                    if (quanti_dipendenti < result.Count)
                    {
                        wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                Thread.Sleep(1000);

                return new_result;
            }
        }
        public static List<Dictionary<string, string>> cattura_amministratori()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            List<string> amm = new List<string>();
            List<List<string>> result = new List<List<string>>();
            List<Dictionary<string, string>> new_result = new List<Dictionary<string, string>>();

            while (true)
            {
                Thread.Sleep(1000);
                while (true)
                {
                    try
                    {
                        wait.Until(d => d.FindElement(By.LinkText("Amministrazione"))).Click();
                        Thread.Sleep(1000);
                        break;
                    }
                    catch (ElementClickInterceptedException)
                    {
                        new Actions(driver).SendKeys(Keys.End).Perform();
                        continue;
                    }
                    catch (Exception)
                    {
                        return new_result;
                    }
                }

                if (HTML.vedisece("AMMINISTRAZIONE", driver))
                {
                    break;
                }
                else
                {
                    wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                    Thread.Sleep(1000);
                    new Actions(driver).SendKeys(Keys.End).Perform();
                    continue;
                }
            }

            while (true)
            {
                IWebElement lista_dipen = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'amministrazione')]")));
                foreach (IWebElement li in lista_dipen.FindElements(By.XPath(".//li")))
                {
                    List<string> tmp = li.FindElement(By.Id("dettaglio_iscritto")).Text.Split('\n').ToList();

                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    foreach (string s in tmp)
                    {
                        string descr = s.Substring(0, s.IndexOf(":"));
                        string valor = string.Empty;
                        string x = string.Empty;

                        if (descr.Contains("|"))
                        {
                            valor = s.Substring(0, s.IndexOf("|")).Trim();
                            descr = "NOMINATIVO";

                            dict.Add(descr, valor);

                            x = s.Substring(s.IndexOf("|") + 1).Trim();
                            descr = x.Substring(0, x.IndexOf(":"));
                            valor = x.Substring(x.IndexOf(":") + 1);

                            dict.Add(descr, valor);
                        }
                        else
                        {
                            switch (descr)
                            {
                                case "LUOGO DI NASCITA":
                                    valor = s.Substring(s.IndexOf(":") + 1);
                                    valor = valor.Substring(0, valor.IndexOf(" DATA DI"));
                                    dict.Add(descr, valor);

                                    x = s.Substring(s.IndexOf(valor) + valor.Length).Trim();
                                    descr = x.Substring(x.IndexOf("DATA DI"), x.IndexOf(":"));
                                    valor = x.Substring(x.IndexOf(":") + 1).Trim();
                                    dict.Add(descr, valor);

                                    break;

                                case "SESSO":
                                    valor = s.Substring(s.IndexOf(":") + 1);
                                    valor = valor.Substring(0, valor.IndexOf(" CODICE"));
                                    dict.Add(descr, valor);

                                    x = s.Substring(s.IndexOf(valor) + valor.Length).Trim();
                                    descr = x.Substring(x.IndexOf("CODICE"), x.IndexOf(":"));
                                    valor = x.Substring(x.IndexOf(":") + 1).Trim();
                                    dict.Add(descr, valor);

                                    break;

                                default:
                                    valor = s.Substring(s.IndexOf(':') + 1).Trim();
                                    dict.Add(descr, valor);

                                    break;
                            }
                        }
                    }
                    new_result.Add(dict);
                }
                if (new_result.Count == 0)
                {
                    continue;
                }
                else
                {
                    wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                    Thread.Sleep(1000);
                    return new_result;
                }
            }
        }
        public static List<Dictionary<string, string>> cattura_storico()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            List<string> stor = new List<string>();
            List<List<string>> result = new List<List<string>>();
            List<Dictionary<string, string>> new_result = new List<Dictionary<string, string>>();

            while (true)
            {
                Thread.Sleep(1000);
                while (true)
                {
                    try
                    {
                        wait.Until(d => d.FindElement(By.LinkText("Storico Stati"))).Click();
                        Thread.Sleep(1000);
                        break;
                    }
                    catch (ElementClickInterceptedException)
                    {
                        new Actions(driver).SendKeys(Keys.End).Perform();
                        continue;
                    }
                    catch (Exception)
                    {
                        return new_result;
                    }
                }

                if (HTML.vedisece("STORICO STATI", driver))
                {
                    break;
                }
                else
                {
                    wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                    Thread.Sleep(1000);
                    new Actions(driver).SendKeys(Keys.End).Perform();
                    continue;
                }
            }

            while (true)
            {
                IWebElement lista_dipen = wait.Until(d => d.FindElement(By.XPath("//ul[contains(@class, 'storico')]")));
                foreach (IWebElement li in lista_dipen.FindElements(By.XPath(".//li")))
                {
                    List<string> tmp = li.Text.Split('\n').ToList();

                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    foreach (string s in tmp)
                    {
                        string descr = s.Substring(0, s.IndexOf(":")).Trim();
                        string valor = s.Substring(s.IndexOf(":") + 1).Trim();


                        dict.Add(descr, valor);
                    }
                    new_result.Add(dict);
                }

                if (new_result.Count == 0)
                {
                    continue;
                }
                else if (new_result.Count > 0)
                {
                    if (new_result[0].Count < 4)
                    {
                        continue;
                    }
                }
                wait.Until(d => d.FindElement(By.LinkText("Indietro"))).Click();
                Thread.Sleep(1000);
                return new_result;
            }
        }
        public static void collega_driver(ChromeDriver browser)
        {
            driver = browser;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
    }
    public partial class Funzioni : ResourceDictionary
    {
        #region COMBOBOX & DATEPICKER
        /// <summary>
        /// Variabile per eseguire una sola volta l'evento SelectionChanged
        /// </summary>
        private bool boolChange = false;
        private bool filtra = true;
        private bool Click = false;

        private List<ComboBoxItem> save_lista = [];

        private Dictionary<string, int> mesi = new Dictionary<string, int>()
        {
            { "gennaio", 31 },
            { "febbraio", DateTime.Now.Year % 4 == 0 ? 29 : 28 },
            { "marzo", 31 },
            { "aprile", 30 },
            { "maggio", 31 },
            { "giugno", 30 },
            { "luglio", 31 },
            { "agosto", 31 },
            { "settembre", 30 },
            { "ottobre", 31 },
            { "novembre", 30 },
            { "dicembre", 31 }
        };

        private void load_listaGiorni(object obj, RoutedEventArgs e)
        {
            int giorno_corrente = int.Parse(DateTime.Now.ToString("dd"));
            string mese_corrente = DateTime.Now.ToString("MMMM");
            int giorni_mese = this.mesi[mese_corrente];

            ListBox lista = obj as ListBox;

            if (lista.Items.Count == 0)
            {
                List<ListBoxItem> lista_giorni = new List<ListBoxItem>();
                for (int z = 0; z < 2; z++)
                {
                    for (int giorno = 1; giorno <= giorni_mese; giorno++)
                    {
                        lista_giorni.Add(new ListBoxItem()
                        {
                            Content = giorno.ToString(),
                            HorizontalContentAlignment = HorizontalAlignment.Center,
                            VerticalContentAlignment = VerticalAlignment.Center,
                            Height = imposta_altezza(lista)
                        });
                    }
                }

                lista.ItemsSource = lista_giorni;
                lista.SelectedIndex = 16; //this.lista_giorni.Count / 2 + (giorno_corrente - 1);
            }
            else
            {

                int n = Elementi_visibili(lista);

                ScrollViewer scroll = GetScrollViewer(obj as ListBox);
                scroll.ScrollToVerticalOffset(lista.SelectedIndex - (n / 2));
            }
        }
        private void load_listaMesi(object obj, RoutedEventArgs e)
        {
            string mese_corrente = DateTime.Now.ToString("MMMM");

            ListBox lista = obj as ListBox;
            if (lista.Items.Count == 0)
            {
                List<ListBoxItem> lista_mesi = new List<ListBoxItem>();
                for (int i = 0; i < 2; i++)
                {
                    foreach (string mese in this.mesi.Keys)
                    {
                        lista_mesi.Add(new ListBoxItem()
                        {
                            Content = mese.Substring(0, 1).ToUpper() + mese.Substring(1),
                            HorizontalContentAlignment = HorizontalAlignment.Center,
                            VerticalContentAlignment = VerticalAlignment.Center,
                            Padding = new Thickness(0),
                            Height = imposta_altezza(lista)
                        });
                    }
                }

                lista.ItemsSource = lista_mesi;
                lista.SelectedIndex = (this.mesi.Keys.ToList().IndexOf(mese_corrente));
            }
            else
            {
                int n = Elementi_visibili(lista);

                ScrollViewer scroll = GetScrollViewer(lista);
                scroll.ScrollToVerticalOffset(lista.SelectedIndex - (n / 2));
            }
        }
        private void load_listaAnni(object obj, RoutedEventArgs e)
        {
            int anno = DateTime.Now.Year;
            ListBox lista = obj as ListBox;
            ScrollViewer scroll = GetScrollViewer(lista);
            List<ListBoxItem> lista_anni = new List<ListBoxItem>();

            if (lista.Items.Count == 0)
            {
                for (int i = 1900; i <= anno + 30; i++)
                {
                    lista_anni.Add(new ListBoxItem()
                    {
                        Content = i.ToString(),
                        HorizontalContentAlignment = HorizontalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        Height = imposta_altezza(lista)
                    });
                }
                lista.ItemsSource = lista_anni;
            }
            else
            {
                lista_anni = lista.ItemsSource as List<ListBoxItem>;
                int n = Elementi_visibili(lista);

                if (anno - 1900 + 30 == lista.Items.Count - 1)
                {
                    for (int i = 0; i < (n / 2); i++)
                    {
                        lista_anni.Insert(0,
                            new ListBoxItem()
                            {
                                Content = " ",
                                IsHitTestVisible = false,
                                Height = imposta_altezza(lista)
                            });

                        lista_anni.Add(new ListBoxItem()
                        {
                            Content = " ",
                            IsHitTestVisible = false,
                            Height = imposta_altezza(lista)
                        });
                    }

                    lista.ItemsSource = null;
                    lista.ItemsSource = lista_anni;

                    lista.ApplyTemplate();
                    lista.SelectedIndex = anno - 1900 + (n / 2);
                }

                scroll.ScrollToVerticalOffset(lista.SelectedIndex - (n / 2));
            }
        }

        /// <summary>
        /// Restituisce il numero di elementi di una ListBox visibili nello ScrollViewer
        /// </summary>
        /// <param name="lista"></param>
        /// <returns></returns>
        private int Elementi_visibili(ListBox lista)
        {
            Popup popup = FindParent(lista, typeof(Popup));
            lista.ApplyTemplate();
            lista.UpdateLayout();

            return (int)Math.Floor((popup.Height - 35) / (lista.Items.GetItemAt(0) as ListBoxItem).Height);
        }
        private int imposta_altezza(object elem)
        {
            Popup popup = FindParent(elem, typeof(Popup));
            return (int)popup.Height / (9 + 1);
        }
        private ScrollViewer GetScrollViewer(DependencyObject depObj, bool figli = true)
        {
            if (depObj is ScrollViewer) return depObj as ScrollViewer;

            if (figli)
            {
                for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    var child = System.Windows.Media.VisualTreeHelper.GetChild(depObj, i);
                    var result = GetScrollViewer(child);
                    if (result != null) return result;
                }
            }
            else
            {
                var parent = System.Windows.Media.VisualTreeHelper.GetParent(depObj);
                var result = GetScrollViewer(parent, false);
                if (result != null) return result;
            }
            return null;
        }
        /// <summary>
        /// Apertura Popup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ApriDropDown(object sender, RoutedEventArgs e)
        {
            dynamic tmp = sender;     // COMBOBOX || DATEPICKER
            tmp.IsDropDownOpen = true;
        }
        /// <summary>
        /// Imposta il fuoco all'interno del Popup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void focus(object sender, EventArgs e)
        {
            (((sender as Popup).Child as Border).Child as Grid).Children[0].Focus();
        }
        /// <summary>
        /// Evento scatenato alla pressione del tasto del mouse
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void list_ClickDown(object sender, RoutedEventArgs e)
        {
            if ((e.OriginalSource as dynamic).Name != "thumb")
            {
                this.Click = true;
            }
        }
        /// <summary>
        /// Chiude il Popup dopo la selezione di un elemento nella ComboBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>   
        /// 
        /*
        private void selezione(object sender, RoutedEventArgs e)
        {
            if (this.Click)
            {
                this.Click = false;
                if (e.OriginalSource.GetType() != typeof(ScrollViewer))
                {
                    var clickedElem = (e.OriginalSource as dynamic).TemplatedParent;
                    if (clickedElem.GetType() != typeof(ScrollBar))
                    {
                        dynamic tmp = (sender as ListBox).TemplatedParent;
                        tmp.IsDropDownOpen = false;

                        IconTextBox tbox = tmp.Template.FindName("tbox", tmp) as IconTextBox;

                        Filtra(tbox, null);
                        tbox.Focus();
                    }
                }
            }
            this.mouse_pressed = false;
        }
        */
        /// <summary>
        /// Evento attivato allo scroll nella ListBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scroll(object sender, MouseWheelEventArgs e)
        {
            ListBox list = sender as ListBox;
            ScrollViewer scroll = GetScrollViewer(list);

            int tmp = 0;
            if (scroll.VerticalOffset == scroll.ScrollableHeight || scroll.VerticalOffset == 0) // FINE LISTA
            {
                if (list.Name != "lista_anni")
                {
                    tmp = list.SelectedIndex;
                    list.SelectedIndex = tmp;
                }
            }

            this.boolChange = true;
            if (e.Delta < 0)
            {
                if (list.Name == "lista_anni")
                { // SE LISTA_ANNI NON SELEZIONO I LISTBOXITEM VUOTI IN FONDO ALLA LISTA
                    if (list.SelectedIndex < list.Items.Count - 5)
                    {
                        list.SelectedIndex += 1;
                    }
                }
                else
                {
                    list.SelectedIndex += 1;
                }
            }
            else if (list.SelectedIndex > 0)
            {
                if (list.Name == "lista_anni")
                { // SE LISTA_ANNI NON SELEZIONO I LISTBOXITEM VUOTI ALL'INIZIO DELLA LISTA
                    if (list.SelectedIndex >= 5)
                    {
                        list.SelectedIndex -= 1;
                    }
                }
                else
                {
                    list.SelectedIndex -= 1;
                }
            }

            e.Handled = true;
        }
        private void selection(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && e.RemovedItems.Count > 0)
            {
                ListBox lista = sender as ListBox;
                int index = lista.Items.IndexOf(e.AddedItems[0]);
                int old_index = lista.Items.IndexOf(e.RemovedItems[0]);

                int index_diff = index - old_index;

                if (index_diff + (lista.Items.Count / 2) == 0)
                {
                    return;
                }

                if (this.boolChange)
                {
                    this.boolChange = false;
                }
                else
                {
                    this.boolChange = true;
                    return;
                }

                ScrollViewer scroll = GetScrollViewer(lista);
                if (index_diff > 0)
                {
                    if (scroll.VerticalOffset + index_diff > scroll.ScrollableHeight)
                    {
                        this.boolChange = true;
                        lista.SelectedIndex = lista.SelectedIndex - (lista.Items.Count / 2);
                        scroll.ScrollToVerticalOffset(lista.SelectedIndex - 4); // ----------------------------------------------------------------------- elementi visibili
                        this.boolChange = false;
                    }
                    else
                    {
                        for (int i = index_diff; i > 0; i--)
                        {
                            scroll.LineDown();
                        }
                    }
                }
                else if (index_diff < 0)
                {
                    if (scroll.VerticalOffset + index_diff <= 0 && lista.Name != "lista_anni")
                    {
                        this.boolChange = true;
                        lista.SelectedIndex = (lista.Items.Count / 2) + lista.SelectedIndex;
                        scroll.ScrollToVerticalOffset(lista.SelectedIndex - 4); // ----------------------------------------------------------------------- elementi visibili
                        this.boolChange = false;
                    }
                    else
                    {
                        for (int i = index_diff; i < 0; i++)
                        {
                            scroll.LineUp();
                        }
                    }
                }
                lista.UpdateLayout();

                string valore_selezionato = ((ListBoxItem)lista.SelectedValue).Content.ToString().ToLower();

                DatePicker tmp = lista.TemplatedParent as DatePicker;
                ListBox listBox_giorni = tmp.Template.FindName("lista_giorni", tmp) as ListBox;
                ListBox listBox_mesi = tmp.Template.FindName("lista_mesi", tmp) as ListBox;

                ScrollViewer scroll_giorni = GetScrollViewer(listBox_giorni);
                int giorno_selezionato = listBox_giorni.SelectedIndex + 1;
                int mese_selezionato = listBox_mesi.SelectedIndex + 1;

                List<ListBoxItem> lista_giorni = listBox_giorni.ItemsSource as List<ListBoxItem>;
                switch (lista.Name)
                {
                    case "lista_mesi":
                        if (this.mesi.ContainsKey(valore_selezionato))
                        {
                            int giorni_mese = this.mesi[valore_selezionato];

                            if (listBox_giorni.Items.Count != giorni_mese * 2)
                            {
                                lista_giorni = new List<ListBoxItem>();
                                for (int z = 0; z < 2; z++)
                                {
                                    for (int giorno = 1; giorno <= giorni_mese; giorno++)
                                    {
                                        lista_giorni.Add(new ListBoxItem()
                                        {
                                            Content = giorno.ToString(),
                                            HorizontalContentAlignment = HorizontalAlignment.Center,
                                            VerticalContentAlignment = VerticalAlignment.Center,
                                            Height = imposta_altezza(lista)
                                        });
                                    }
                                }
                            }
                        }
                        break;

                    case "lista_anni":
                        if (int.Parse(valore_selezionato) % 4 == 0)
                        {
                            this.mesi["febbraio"] = 29;
                        }
                        else
                        {
                            this.mesi["febbraio"] = 28;
                        }

                        if (mese_selezionato == 2 || mese_selezionato == 14) // SE HO SELEZIONATO FEBBRAIO
                        {
                            int giorni_mese = this.mesi["febbraio"];
                            if (listBox_giorni.Items.Count != giorni_mese * 2)
                            {
                                lista_giorni = new List<ListBoxItem>();
                                for (int z = 0; z < 2; z++)
                                {
                                    for (int giorno = 1; giorno <= giorni_mese; giorno++)
                                    {
                                        lista_giorni.Add(new ListBoxItem()
                                        {
                                            Content = giorno.ToString(),
                                            HorizontalContentAlignment = HorizontalAlignment.Center,
                                            VerticalContentAlignment = VerticalAlignment.Center,
                                            Height = imposta_altezza(lista)
                                        });
                                    }
                                }
                            }
                        }
                        else
                        {
                            lista_giorni = listBox_giorni.ItemsSource as List<ListBoxItem>;
                        }
                        break;
                }

                if (giorno_selezionato > (listBox_giorni.Items.Count / 2))
                {
                    giorno_selezionato -= (listBox_giorni.Items.Count / 2);
                }
                else if (giorno_selezionato < (lista_giorni.Count / 2))
                {
                    giorno_selezionato -= (lista_giorni.Count / 2);
                }
                else if (giorno_selezionato >= (lista_giorni.Count / 2))
                {
                    giorno_selezionato = 0;
                }
                else
                { // INDICE SULL'ULTIMO ELEMENTO DELLA LISTA
                    giorno_selezionato = (listBox_giorni.Items.Count - lista_giorni.Count) / 2;
                }

                int save_index = giorno_selezionato + (lista_giorni.Count / 2) - 1;

                listBox_giorni.ItemsSource = lista_giorni;
                listBox_giorni.SelectedIndex = save_index;
                listBox_giorni.Items.Refresh();

                int n = Elementi_visibili(lista);
                scroll_giorni.ScrollToVerticalOffset(listBox_giorni.SelectedIndex - (n / 2));
            }
        }
        /// <summary>
        /// Azione del tasto nella ListBox per salire nello ScrollViewer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sali(object sender, EventArgs e)
        {
            ListBox list = (sender as RepeatButton).TemplatedParent as ListBox;
            //ListBox list = (sender as Button).TemplatedParent as ListBox;
            ScrollViewer scroll = GetScrollViewer(list);

            this.boolChange = true;
            if (list.Name == "lista_anni")
            { // SE LISTA_ANNI NON SELEZIONO I LISTBOXITEM VUOTI
                if (list.SelectedIndex > 4)
                {
                    list.SelectedIndex -= 1;
                }
            }
            else
            {
                if (scroll.VerticalOffset == 0)
                {
                    this.boolChange = false;
                    list.SelectedIndex = (list.Items.Count / 2) + list.SelectedIndex;
                }
                this.boolChange = true;
                list.SelectedIndex -= 1;
            }
        }
        /// <summary>
        /// Azione del tasto nella ListBox per scendere nello ScrollViewer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scendi(object sender, EventArgs e)
        {
            //ListBox list = (sender as Button).TemplatedParent as ListBox;
            ListBox list = (sender as RepeatButton).TemplatedParent as ListBox;
            ScrollViewer scroll = GetScrollViewer(list);

            this.boolChange = true;
            if (list.Name == "lista_anni")
            {
                if (list.SelectedIndex < list.Items.Count - 5)
                {
                    list.SelectedIndex += 1;
                }
            }
            else
            {
                if (scroll.VerticalOffset == scroll.ScrollableHeight)
                {
                    this.boolChange = false;
                    int tmp = list.SelectedIndex;
                    list.SelectedIndex = tmp;
                }
                this.boolChange = true;
                list.SelectedIndex += 1;
            }
        }
        /*
        private void Filtra(object sender, KeyEventArgs e)
        {
            IconTextBox tbox = sender as IconTextBox;
            if (tbox.Search)
            {
                string testo = tbox.Text;
                TextBox t = tbox.Template.FindName("tbox", tbox) as TextBox;

                ComboBox combo = tbox.TemplatedParent as ComboBox;
                ListBox lista = combo.Template.FindName("lista", combo) as ListBox;
                BindingExpression bind_exp = lista.GetBindingExpression(ListBox.ItemsSourceProperty);
                Binding bind = bind_exp.ParentBinding;

                ObservableCollection<ComboBoxItem> filtered_list = [];
                foreach (ComboBoxItem elem in Extensions.GetComboBoxSaveList(combo))
                {
                    if (elem.Content.ToString().StartsWith(t.Text.ToUpper()))
                    {
                        filtered_list.Add(new ComboBoxItem() { Content = elem.Content.ToString() });
                    }
                }
                combo.ItemsSource = filtered_list;
                combo.Items.Refresh();

                lista.Dispatcher.Invoke(() => lista.ItemsSource = filtered_list);
                lista.SetBinding(ListBox.ItemsSourceProperty, bind);

                tbox.Text = testo;
                combo.Text = testo;
            }
        }
        */
        private void ok(object sender, MouseButtonEventArgs e)
        {
            string ret = string.Empty;

            DatePicker datePicker = (sender as Button).TemplatedParent as DatePicker;
            ListBox lista_giorni = datePicker.Template.FindName("lista_giorni", datePicker) as ListBox;
            ListBox lista_mesi = datePicker.Template.FindName("lista_mesi", datePicker) as ListBox;
            ListBox lista_anni = datePicker.Template.FindName("lista_anni", datePicker) as ListBox;

            ret += (lista_giorni.SelectedItem as ListBoxItem).Content.ToString();
            ret += " ";
            ret += (lista_mesi.SelectedItem as ListBoxItem).Content.ToString();
            ret += " ";
            ret += (lista_anni.SelectedItem as ListBoxItem).Content.ToString();

            datePicker.SelectedDate = DateTime.Parse(ret);
            datePicker.IsDropDownOpen = false;
        }
        private void chiudi(object sender, EventArgs e)
        {
            DatePicker tmp = (sender as dynamic).TemplatedParent as DatePicker;
            tmp.IsDropDownOpen = false;
        }
        #endregion

        #region SCROLLBAR

        private bool mouse_pressed = false;
        private bool mouse_leave = true;

        private Point startPosition = new Point(0, 0);

        private Border thumb = null;

        /// <summary>
        /// Click barra scorrimento
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scroll_click(object sender, MouseButtonEventArgs e)
        {
            ScrollBar scroll = (sender as Grid).TemplatedParent as ScrollBar;
            ScrollViewer view = GetScrollViewer(scroll, false);

            Border thumb = scroll.Template.FindName("thumb", scroll) as Border;

            Point p = e.GetPosition(thumb);
            if (p.Y < 0)
            {
                view.PageUp();
            }
            else
            {
                view.PageDown();
            }
        }
        /// <summary>
        /// Click pulsanti barra scorrimento
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scroll_btn(object sender, RoutedEventArgs e)
        {
            RepeatButton button = sender as RepeatButton;
            ScrollBar scroll = button.TemplatedParent as ScrollBar;
            ScrollViewer view = GetScrollViewer(scroll, false);

            if (button.Name == "btn_up")
            {
                view.LineUp();
            }
            else if (button.Name == "btn_down")
            {
                view.LineDown();
            }
        }
        /// <summary>
        /// Scroll da trascinamento thumb
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scroll_drag(object sender, MouseEventArgs e)
        {
            if (this.mouse_pressed)
            {
                ScrollBar scroll = this.thumb.TemplatedParent as ScrollBar;
                ScrollViewer view = GetScrollViewer(scroll, false);

                Point p = e.GetPosition(this.thumb);

                if (p.X >= this.startPosition.X - 200 && p.X <= this.startPosition.X + 200)
                {
                    if (p.Y > startPosition.Y + 5)
                    {
                        view.LineDown();
                    }
                    else if (p.Y < startPosition.Y - 5)
                    {
                        view.LineUp();
                    }
                }
            }
        }
        /// <summary>
        /// Rilascio click mouse
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scroll_drag_up(object sender, MouseEventArgs e)
        {
            ScrollBar scroll = this.thumb.TemplatedParent as ScrollBar;
            ScrollViewer view = GetScrollViewer(scroll, false);

            view.MouseMove -= scroll_drag;
            this.mouse_pressed = false;
        }
        /// <summary>
        /// Click mouse
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scroll_drag_down(object sender, MouseEventArgs e)
        {
            this.thumb = sender as Border;

            ScrollBar scroll = this.thumb.TemplatedParent as ScrollBar;
            ScrollViewer view = GetScrollViewer(scroll, false);

            view.MouseMove += scroll_drag;

            this.mouse_pressed = true;
            this.startPosition = e.GetPosition(this.thumb);
        }
        #endregion
        /// <summary>
        /// Trova e restituisce il primo contenitore (del tipo specificato) dell'elemento
        /// </summary>
        /// <param name="elemento">Elemento di cui trovare il contenitore</param>
        /// <param name="tipo">Tipo del contenitore da ricercare</param>
        /// <returns></returns>
        private dynamic FindParent(dynamic elemento, Type tipo)
        {
            object elem = elemento.Parent;

            if (elem.GetType() == tipo)
            {
                return elem;
            }
            else if (elem != null)
            {
                return FindParent(elem, tipo);
            }
            return null;
        }


        private void FocusTextBox(object sender, RoutedEventArgs e)
        {
            dynamic elem = sender as TextBox;
            if (elem == null)
            {
                elem = sender as ComboBox;
            }

            if (e.OriginalSource == elem)
            {
                (elem.Template.FindName("tbox", elem) as UIElement).Focus();
            }
        }
        public bool click_down = false;
        private void ComboBoxKey(object sender, KeyEventArgs e)
        {
            if (e.OriginalSource.GetType() == typeof(ListBox) ||
                e.OriginalSource.GetType() == typeof(ListBoxItem))
            {
                return;
            }

            ListBox lista;
            switch (e.Key)
            {
                case Key.Enter:
                    if (e.IsDown)
                    {
                        this.click_down = true;
                        break;
                    }
                    if (click_down)
                    {
                        (sender as ComboBox).IsDropDownOpen = true;
                        this.click_down = false;
                    }
                    break;

                case Key.Down:
                    if (e.IsUp)
                    {
                        lista = (sender as ComboBox).Template.FindName("lista", (sender as ComboBox)) as ListBox;

                        if (lista.SelectedIndex < lista.Items.Count)
                        {
                            lista.SelectedIndex += 1;
                        }
                    }
                    break;

                case Key.Up:
                    if (e.IsUp)
                    {
                        lista = (sender as ComboBox).Template.FindName("lista", (sender as ComboBox)) as ListBox;
                        if (lista.SelectedIndex > 0)
                        {
                            lista.SelectedIndex -= 1;
                        }
                    }
                    break;
            }

        }
    }
    public static class Extensions
    {
        #region STANDARD
        public static readonly DependencyProperty StandardForegroundProperty =
            DependencyProperty.RegisterAttached("StandardForeground", typeof(System.Windows.Media.SolidColorBrush), typeof(Extensions),
                new PropertyMetadata(new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Black)));
        public static System.Windows.Media.SolidColorBrush GetStandardForeground(DependencyObject element)
        {
            return (System.Windows.Media.SolidColorBrush)element.GetValue(StandardForegroundProperty);
        }
        public static void SetStandardForeground(DependencyObject element, System.Windows.Media.SolidColorBrush value)
        {
            element.SetValue(StandardForegroundProperty, value);
        }

        public static readonly DependencyProperty StandardBackgroundProperty =
            DependencyProperty.RegisterAttached("StandardBackground", typeof(System.Windows.Media.SolidColorBrush), typeof(Extensions),
                new PropertyMetadata(new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White)));
        public static System.Windows.Media.SolidColorBrush GetStandardBackground(DependencyObject element)
        {
            return (System.Windows.Media.SolidColorBrush)element.GetValue(StandardBackgroundProperty);
        }
        public static void SetStandardBackground(DependencyObject element, System.Windows.Media.SolidColorBrush value)
        {
            element.SetValue(StandardBackgroundProperty, value);
        }
        #endregion

        #region HOVER
        public static readonly DependencyProperty HoverForegroundProperty =
            DependencyProperty.RegisterAttached("HoverForeground", typeof(System.Windows.Media.SolidColorBrush), typeof(Extensions),
                new PropertyMetadata(new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Black)));
        public static System.Windows.Media.SolidColorBrush GetHoverForeground(DependencyObject element)
        {
            return (System.Windows.Media.SolidColorBrush)element.GetValue(HoverForegroundProperty);
        }
        public static void SetHoverForeground(DependencyObject element, System.Windows.Media.SolidColorBrush value)
        {
            element.SetValue(HoverForegroundProperty, value);
        }

        public static readonly DependencyProperty HoverBackgroundProperty =
            DependencyProperty.RegisterAttached("HoverBackground", typeof(System.Windows.Media.SolidColorBrush), typeof(Extensions),
                new PropertyMetadata(new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightGray)));
        public static System.Windows.Media.SolidColorBrush GetHoverBackground(DependencyObject element)
        {
            return (System.Windows.Media.SolidColorBrush)element.GetValue(HoverForegroundProperty);
        }
        public static void SetHoverBackground(DependencyObject element, System.Windows.Media.SolidColorBrush value)
        {
            element.SetValue(HoverForegroundProperty, value);
        }
        #endregion

        #region PRESSED
        public static readonly DependencyProperty PressedForegroundProperty =
            DependencyProperty.RegisterAttached("PressedForeground", typeof(System.Windows.Media.SolidColorBrush), typeof(Extensions),
                new PropertyMetadata(new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Black)));
        public static System.Windows.Media.SolidColorBrush GetPressedForeground(DependencyObject element)
        {
            return (System.Windows.Media.SolidColorBrush)element.GetValue(PressedForegroundProperty);
        }
        public static void SetPressedForeground(DependencyObject element, System.Windows.Media.SolidColorBrush value)
        {
            element.SetValue(PressedForegroundProperty, value);
        }

        public static readonly DependencyProperty PressedBackgroundProperty =
            DependencyProperty.RegisterAttached("PressedBackground", typeof(System.Windows.Media.SolidColorBrush), typeof(Extensions),
        new PropertyMetadata(new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.CadetBlue)));
        public static System.Windows.Media.SolidColorBrush GetPressedBackground(DependencyObject element)
        {
            return (System.Windows.Media.SolidColorBrush)element.GetValue(PressedBackgroundProperty);
        }
        public static void SetPressedBackground(DependencyObject element, System.Windows.Media.SolidColorBrush value)
        {
            element.SetValue(PressedBackgroundProperty, value);
        }
        #endregion

        #region COMBOBOX SAVE LIST
        public static readonly DependencyProperty ComboBoxSaveListProperty =
                        DependencyProperty.RegisterAttached("ComboBoxSaveList", typeof(List<ComboBoxItem>), typeof(Extensions),
                            new PropertyMetadata(new List<ComboBoxItem>()));
        public static List<ComboBoxItem> GetComboBoxSaveList(DependencyObject element)
        {
            return element.GetValue(ComboBoxSaveListProperty) as List<ComboBoxItem>;
        }
        public static void SetComboBoxSaveList(DependencyObject element, List<ComboBoxItem> value)
        {
            element.SetValue(ComboBoxSaveListProperty, value);
        }
        #endregion
    }

    #region CONVERTITORI
    public class SizeCheckBox : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (int.Parse(value.ToString()) * 1.6);
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (int.Parse(value.ToString()) / 1.6);
        }
    }

    public class InvertBool : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (!(bool)value);
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return ((bool)value);
        }
    }

    public class ConvertDate : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is null ? DateTime.Now.ToString("dd MMMM yyyy") : ((DateTime)value).ToString("dd MMMM yyyy");
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DateTime.Parse((string)value);
        }
    }

    public class ComboBoxItems2ListBoxItem : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            List<ListBoxItem> ret = new List<ListBoxItem>();
            foreach (ComboBoxItem item in value as ItemCollection)
            {
                ret.Add(new ListBoxItem() { Content = item.Content.ToString() });
            }
            return ret;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            List<ComboBoxItem> ret = new List<ComboBoxItem>();
            foreach (ListBoxItem item in value as ItemCollection)
            {
                ret.Add(new ComboBoxItem() { Content = item.Content.ToString() });
            }
            return ret;
        }
    }

    public class CreaMargin : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, CultureInfo culture)
        {
            //  237 : 85 = x : value
            // (237 / 85) * value = x

            double n = (258.00 - int.Parse(value[2].ToString())) / (double)value[1];
            double top_margin = Math.Round(n * int.Parse(value[0].ToString()));

            return new Thickness(0, top_margin, 0, 0);
        }
        public object[] ConvertBack(object value, Type[] targetType, object parameter, CultureInfo culture)
        {
            //  85 : 237 = x : value
            // (85 / 237) * value = x
            throw new NotImplementedException();
        }
    }

    public class AltezzaThumb : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, CultureInfo culture)
        {
            // value - 46 (40 altezza_pulsanti / 6 altezza margin)
            double spazio = double.Parse(value[0].ToString()) - 40;
            int passo = int.Parse(value[1].ToString()) / 5;

            if (spazio > 0)
                return spazio - (5 * passo);
            else
                return 15;
        }
        public object[] ConvertBack(object value, Type[] targetType, object parameter, CultureInfo culture)
        {
            //  85 : 237 = x : value
            // (85 / 237) * value = x
            throw new NotImplementedException();
        }
    }
    
    #endregion

}

