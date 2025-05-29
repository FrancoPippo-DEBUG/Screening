using DLL;
using OpenQA.Selenium.Chrome;
using PdfSharp.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace Avvera
{
    internal class Avvera
    {
        private static PDF pdf = new PDF();
        private static XFont bold = new XFont("Tahoma", 14, XFontStyleEx.Bold);
        private static XFont mini_bold = new XFont("Tahoma", 10, XFontStyleEx.Bold);

        [STAThread]
        static void Main(string[] args)
        {
            Win win = new Win("Avvera");
            ChromeDriver? browser = null;

            LogFile log = new LogFile();
            Excel excel = new Excel();

            pdf.pos_x = XUnit.FromPoint(40);

            pdf.set_font("Tahoma", 10);
            pdf.spacing = XUnit.FromPoint(15);

            string[] riga_letta = new string[] { };
            string[] natura_giuridica = new string[] { "SRLS", "S.R.L.S", " SRL", " S.R.L.", " SAS", " S.A.S.", " SNC", " S.N.C", " SPA", " S.P.A." };

            List<Dictionary<string, string>> storico = new List<Dictionary<string, string>>();
            List<Dictionary<string, string>> intermediari = new List<Dictionary<string, string>>();
            List<List<string>> prodotti = new List<List<string>>();

            List<string> riga_out = new List<string>();

            int riga = 0;
            int i = 0;
            int indice_intermed = 0;
            int indice_prod = 0;
            try
            {
                string tmp = new InputBox("Riga inizio").result;
                if (tmp == "")
                {
                    riga = 3;
                }
                else
                {
                    riga = int.Parse(tmp);
                }
            }
            catch (Exception)
            {
                Environment.Exit(0);
            }
            MessageFrame msg = new MessageFrame(1);

            try
            {
                excel.Open(Utils.CurDir() + "\\Avvera.xlsx");
                Directory.CreateDirectory(Utils.CurDir() + "\\PDF");

                browser = HTML.ApriBrowser("CR");
                while (true)
                {
                    riga_letta = excel.ReadExcel(riga);
                    msg.scrivi("Elaboro riga " + riga);
                    if (riga_letta[0] == "")
                    {
                        break;
                    }

                    string denominazione = riga_letta[0].ToUpper();
                    string cognome = string.Empty; string nome = string.Empty;
                    foreach (char c in denominazione)
                    {
                        switch ((int)c)
                        {
                            case int x when x == 193 ||
                                            x == 192:
                                denominazione = RimpiazzaAccenti(denominazione, "A", x);
                                break;

                            case int x when x == 200 ||
                                            x == 201:
                                denominazione = RimpiazzaAccenti(denominazione, "E", x);
                                break;

                            case int x when x == 204 ||
                                            x == 205:
                                denominazione = RimpiazzaAccenti(denominazione, "I", x);
                                break;

                            case int x when x == 210 ||
                                            x == 211:
                                denominazione = RimpiazzaAccenti(denominazione, "O", x);
                                break;

                            case int x when x == 217 ||
                                            x == 218:
                                denominazione = RimpiazzaAccenti(denominazione, "U", x);
                                break;
                        }
                    }

                    string codice_fiscale = riga_letta[1];
                    if (codice_fiscale.Length > 11)
                    {
                        Utils.risolviCF(codice_fiscale, denominazione, out cognome, out nome);
                    }
                    string data_nascita = string.Empty;
                    if (riga_letta[2] != "") // riga_letta[2]
                    {
                        data_nascita = DateTime.Parse(riga_letta[2], new CultureInfo("it-IT")).ToString("dd/MM/yyyy"); // riga_letta[2]
                    }
                    string denom_soc = riga_letta[3].ToUpper(); // riga_letta[3]
                    if (nome == "")
                    {
                        nome = denom_soc;
                    }

                ivass:
                    #region IVASS
                    browser.Navigate().GoToUrl("https://www.google.it/");
                    browser.Navigate().GoToUrl("https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                    Ivass.collega_driver(browser);

                    pdf.scrivi(denominazione, XBrushes.Black, true, bold);

                    Thread.Sleep(1000);
                    if (File.Exists(Utils.CurDir() + "\\PDF\\" + denominazione + "_" + codice_fiscale + "_RUI.txt"))
                    {
                        File.Delete(Utils.CurDir() + "\\PDF\\" + denominazione + "_" + codice_fiscale + "_RUI.txt");
                    }

                    Dictionary<string, string> dettaglio_ivass = new Dictionary<string, string>();
                    if (codice_fiscale.Length > 11)
                    {
                        dettaglio_ivass = Ivass.ricerca("", DateTime.Parse(data_nascita, new CultureInfo("it-IT")).ToString("dd/MM/yyyy"), nome, cognome);
                    }
                    else
                    {
                        dettaglio_ivass = Ivass.ricerca("", "", nome);
                    }

                    if (dettaglio_ivass.Count > 0)
                    {
                        riga_out.Add("ISCRITTO");
                        riga_out.Add(dettaglio_ivass["NUMERO ISCRIZIONE"]);
                        riga_out.Add(dettaglio_ivass["SEZIONE"]);
                        riga_out.Add("'" + dettaglio_ivass["DATA ISCRIZIONE"]);

                        List<string> lista_intermediari = new List<string>();
                        if (dettaglio_ivass.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                        {
                            lista_intermediari = dettaglio_ivass["INTERMEDIARI PER CUI OPERA"].Split('\n').ToList();
                        }

                        for (int ind_intermed = 0; ind_intermed < 2; ind_intermed++)
                        {
                            if (lista_intermediari.Count > ind_intermed)
                            {
                                riga_out.Add(Utils.sina(lista_intermediari[ind_intermed], "Sezione").Trim());
                            }
                            else
                            {
                                riga_out.Add("");
                            }
                        }

                        pdf.scrivi("Numero Iscrizione: " + dettaglio_ivass["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 10);
                        pdf.scrivi("Sezione: " + dettaglio_ivass["SEZIONE"], XBrushes.Black);
                        if (dettaglio_ivass.ContainsKey("NOMINATIVO"))
                        {
                            pdf.scrivi("Nominativo: " + dettaglio_ivass["NOMINATIVO"], XBrushes.Black);
                        }
                        else
                        {
                            pdf.scrivi("Nominativo: " + dettaglio_ivass["RAGIONE O DENOMINAZIONE SOCIALE"], XBrushes.Black);
                        }
                        pdf.scrivi("Data Iscrizione: " + dettaglio_ivass["DATA ISCRIZIONE"], XBrushes.Black);

                        string tmp = "";
                        if (codice_fiscale.Length > 11)
                        {
                            tmp = "";
                            if (dettaglio_ivass.ContainsKey("LUOGO NASCITA"))
                            {
                                tmp = dettaglio_ivass["LUOGO NASCITA"];
                            }
                            pdf.scrivi("Luogo di Nascita: " + tmp, XBrushes.Black);

                            tmp = "";
                            if (dettaglio_ivass.ContainsKey("DATA NASCITA"))
                            {
                                tmp = dettaglio_ivass["DATA NASCITA"];
                            }
                            pdf.scrivi("Data di Nascita: " + tmp, XBrushes.Black);

                            pdf.scrivi("QUALIFICA DI ESERIZIO", XBrushes.Black, true, bold, 10);
                            List<string> lista_qualifiche = new List<string>();
                            if (dettaglio_ivass.ContainsKey("QUALIFICA DI ESERCIZIO"))
                            {
                                lista_qualifiche = dettaglio_ivass["QUALIFICA DI ESERCIZIO"].Split('\n').ToList();
                            }
                            foreach (string x in lista_qualifiche)
                            {
                                pdf.scrivi(x, XBrushes.Black, false, null, 10);
                            }
                        }
                        else
                        {
                            pdf.scrivi("Sede Legale: " + dettaglio_ivass["SEDE LEGALE"], XBrushes.Black);
                            pdf.scrivi("RESPONSABILI INTERMEDIAZIONE", XBrushes.Black, true, bold, 10);
                            List<string> lista_responsabili = new List<string>();
                            if (dettaglio_ivass.ContainsKey("RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE"))
                            {
                                lista_responsabili = dettaglio_ivass["RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE"].Split('\n').ToList();
                            }

                            foreach (string x in lista_responsabili)
                            {
                                pdf.scrivi(x, XBrushes.Black, false, null, 10);
                            }

                            pdf.scrivi("ADDETTI INTERMEDIAZIONE", XBrushes.Black, true, bold, 10);
                            List<string> lista_addetti_intermed = new List<string>();
                            if (dettaglio_ivass.ContainsKey("ADDETTI ALL'ATTIVITÀ DI INTERMEDIAZIONE"))
                            {
                                lista_addetti_intermed = dettaglio_ivass["ADDETTI ALL'ATTIVITÀ DI INTERMEDIAZIONE"].Split('\n').ToList();
                            }
                            foreach (string x in lista_addetti_intermed)
                            {
                                pdf.scrivi(x, XBrushes.Black, false, null, 10);
                            }
                        }

                        pdf.scrivi("INTERMEDIARI", XBrushes.Black, true, bold, 10);
                        foreach (string x in lista_intermediari)
                        {
                            pdf.scrivi(x, XBrushes.Black, false, null, 10);
                        }
                    }
                    else
                    {
                        pdf.scrivi("NON ISCRITTO", XBrushes.Black, true, bold, 10);

                        riga_out.Add("NON ISCRITTO");
                        riga_out.Add("");
                        riga_out.Add("");
                        riga_out.Add("");
                        riga_out.Add("");
                        riga_out.Add("");
                    }
                    pdf.salva(Utils.CurDir() + @"\PDF\", denominazione + "_" + codice_fiscale + "_RUI.pdf");
                #endregion

                ocf:
                    #region OCF 

                    browser.Navigate().GoToUrl("https://www.organismocf.it/portal/web/portale-ocf/ricerca-nelle-sezioni-dell-albo");
                    OCF.collega_driver(browser);

                    Thread.Sleep(1000);
                    if (File.Exists(Utils.CurDir() + "\\PDF\\" + denominazione + "_" + codice_fiscale + "_OCF.txt"))
                    {
                        File.Delete(Utils.CurDir() + "\\PDF\\" + denominazione + "_" + codice_fiscale + "_OCF.txt");
                    }

                    pdf.scrivi(denominazione, XBrushes.Black, true, bold);
                    Dictionary<string, string> dettaglio_ocf = new Dictionary<string, string>();

                    if (codice_fiscale.Length > 11)
                    {
                        dettaglio_ocf = OCF.cerca(cognome.ToUpper(), nome.ToUpper(), data_nascita);
                    }
                    else
                    {
                        dettaglio_ocf = OCF.cerca(nome.ToUpper(), "", "");
                    }

                    if (dettaglio_ocf.Count > 0)
                    {
                        Thread.Sleep(2000);
                        storico = OCF.storico();

                        riga_out.Add(dettaglio_ocf["STATO ATTUALE"]);

                        pdf.scrivi("DETTAGLIO CONSULENTE", XBrushes.Black, true, bold, 10);
                        pdf.scrivi("Nominativo: " + dettaglio_ocf["NOMINATIVO"], XBrushes.Black, false, null, 10);
                        pdf.scrivi("Sezione Albo: " + dettaglio_ocf["SEZIONE ALBO"], XBrushes.Black);
                        pdf.scrivi("Matricola: " + dettaglio_ocf["MATRICOLA"], XBrushes.Black);

                        if (dettaglio_ocf["STATO ATTUALE"].Contains("CANCELLATO"))
                        {
                            pdf.scrivi("Data di nascita: " + dettaglio_ocf["DATA DI NASCITA"], XBrushes.Black);
                            pdf.scrivi("Luogo di nascita: " + dettaglio_ocf["LUOGO DI NASCITA"], XBrushes.Black);
                            pdf.scrivi("Indirizzo Residenza: " + dettaglio_ocf["INDIRIZZO DEL LUOGO DI CONSERVAZIONE DELLA DOCUMENTAZIONE"], XBrushes.Black);
                            pdf.scrivi("Stato attuale: " + dettaglio_ocf["STATO ATTUALE"], XBrushes.Black);
                        }
                        else
                        {
                            if (dettaglio_ocf.ContainsKey("EMAIL (PEC)"))
                            {
                                pdf.scrivi("PEC: " + dettaglio_ocf["EMAIL (PEC)"], XBrushes.Black);
                            }

                            pdf.scrivi("Data di nascita: " + dettaglio_ocf["DATA DI NASCITA"], XBrushes.Black);
                            pdf.scrivi("Luogo di nascita: " + dettaglio_ocf["LUOGO DI NASCITA"], XBrushes.Black);

                            pdf.scrivi("Indirizzo Domicilio: " + dettaglio_ocf["INDIRIZZO DOMICILIO ELETTO"], XBrushes.Black);
                            pdf.scrivi("Indirizzo Residenza: " + dettaglio_ocf["INDIRIZZO DEL LUOGO DI CONSERVAZIONE DELLA DOCUMENTAZIONE"], XBrushes.Black);
                            pdf.scrivi("Stato attuale: " + dettaglio_ocf["STATO ATTUALE"], XBrushes.Black);
                        }

                        i = 0;
                        pdf.scrivi("STORICO DEGLI STATI", XBrushes.Black, true, bold, 10);
                        foreach (Dictionary<string, string> x in storico)
                        {
                            pdf.entra(4);
                            if (i > 0)
                            {
                                pdf.entra(1);
                                pdf.linea(1);
                            }

                            pdf.scrivi("Stato: " + x["STATO"], XBrushes.Black, false, null, 10);
                            pdf.scrivi("Sezione: " + x["SEZIONE ALBO"], XBrushes.Black);
                            pdf.scrivi("Data delibera: " + x["DATA DELIBERA"], XBrushes.Black, false, null, -1);
                            pdf.scrivi("Data efficacia: " + x["DATA EFFICACIA"], XBrushes.Black, false, null, 0, 380);
                            pdf.scrivi("Numero delibera: " + x["DELIBERA"], XBrushes.Black, false, null, -1);
                            pdf.scrivi("Ente: " + x["ENTE"], XBrushes.Black, false, null, 0, 380);

                            i += 1;
                        }

                        i = 0;
                        intermediari = OCF.intermediari();
                        if (intermediari.Count > 0)
                        {
                            pdf.scrivi("INTERMEDIARI", XBrushes.Black, true, bold, 10);
                        }
                        foreach (Dictionary<string, string> x in intermediari)
                        {
                            pdf.entra(3);
                            if (i > 0)
                            {
                                pdf.entra(1);
                                pdf.linea(1);
                            }
                            pdf.scrivi(x["SOGGETTO"], XBrushes.Black, false, null, 10);
                            pdf.scrivi("Data inizio: " + x["DATA INIZIO"], XBrushes.Black, false, null, -1);
                            pdf.scrivi("Data fine: " + x["DATA FINE"], XBrushes.Black, false, null, 0, 380);
                            i += 1;
                        }

                        if (i > 0)
                        {
                            riga_out.Add(intermediari[0]["SOGGETTO"]);
                        }
                        else
                        {
                            riga_out.Add("");
                        }

                    }
                    else
                    {
                        pdf.scrivi("NON ISCRITTO", XBrushes.Black, true, bold);

                        riga_out.Add("NON ISCRITTO");
                        riga_out.Add("");

                    }
                    pdf.salva(Utils.CurDir() + @"\PDF\", denominazione + "_" + codice_fiscale + "_OCF.pdf");
                #endregion
                
                oam:
                    #region OAM

                    browser.Navigate().GoToUrl("https://www.organismo-am.it/elenchi-registri/filtri.html");
                    OAM.collega_driver(browser);

                    Thread.Sleep(1000);

                    if (File.Exists(Utils.CurDir() + "\\PDF\\" + denominazione + "_" + codice_fiscale + "_OAM.txt"))
                    {
                        File.Delete(Utils.CurDir() + "\\PDF\\" + denominazione + "_" + codice_fiscale + "_OAM.txt");
                    }

                    bool collab = false;
                    bool trovato = OAM.ricerca(codice_fiscale);
                    Dictionary<string, string> dettaglio_oam = new Dictionary<string, string>();
                    if (!trovato)
                    {
                        browser.Navigate().GoToUrl("https://www.organismo-am.it/elenchi-registri/filtri_collaboratori.html");
                        OAM.collega_driver(browser);

                        dettaglio_oam = OAM.ricerca_collab(codice_fiscale);
                        if (dettaglio_oam.Count > 0)
                        {
                            trovato = true;
                            collab = true;
                        }
                    }

                    if (trovato)
                    {
                        Thread.Sleep(2500);

                        if (!collab)
                        {
                            dettaglio_oam = OAM.cattura_dettaglio();

                            riga_out.Add(dettaglio_oam["STATO"]);

                            Dictionary<string, string> sedi_oam = OAM.cattura_sedi();
                            List<Dictionary<string, string>> mandati_diretti_oam = OAM.cattura_mandati("Diretti");
                            List<Dictionary<string, string>> rapp_cessati_diretti = OAM.cattura_mandati("Diretti", true);

                            List<Dictionary<string, string>> mandati_indiretti_oam = OAM.cattura_mandati("Indiretti");
                            List<Dictionary<string, string>> rapp_cessati_indiretti = OAM.cattura_mandati("Indiretti", true);

                            List<Dictionary<string, string>> dipendenti_oam = OAM.cattura_dipendenti();
                            List<Dictionary<string, string>> amministratori_oam = OAM.cattura_amministratori();
                            List<Dictionary<string, string>> storico_oam = OAM.cattura_storico();

                            pdf.scrivi("INFORMAZIONI GENERALI", XBrushes.Black, true, bold);
                            if (codice_fiscale.Length > 11)
                            {
                                pdf.scrivi("Cognome e Nome: " + dettaglio_oam["COGNOME E NOME"], XBrushes.Black, false, null, -1);

                                pdf.scrivi("Cittadinanza: " + (dettaglio_oam["STATO"] != "CANCELLATO" ? dettaglio_oam["CITTADINANZA"] : ""), XBrushes.Black, false, null, 0, 380);

                                pdf.scrivi("Sesso: " + dettaglio_oam["SESSO"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Comune di Nascita: " + dettaglio_oam["COMUNE DI NASCITA"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Codice Fiscale: " + dettaglio_oam["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Provincia di Nascita: " + dettaglio_oam["PROVINCIA DI NASCITA"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Data di Nascita: " + dettaglio_oam["DATA DI NASCITA"], XBrushes.Black, false, null, -1);

                                pdf.scrivi("PEC: " + (dettaglio_oam.ContainsKey("PEC") ? dettaglio_oam["PEC"] : ""), XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Tipo Elenco: " + dettaglio_oam["TIPO ELENCO"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Numero Iscrizione: " + dettaglio_oam["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Stato: " + dettaglio_oam["STATO"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Autorizzato ad Operare: " + dettaglio_oam["AUTORIZZATO AD OPERARE"], XBrushes.Black, false, null, 0, 380);

                                if (dettaglio_oam.ContainsKey("COMUNE"))
                                {
                                    pdf.scrivi("SEDE", XBrushes.Black, true, bold);
                                    pdf.scrivi("Comune: " + dettaglio_oam["COMUNE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Provincia: " + dettaglio_oam["PROVINCIA"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Indirizzo: " + dettaglio_oam["INDIRIZZO"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("CAP: " + dettaglio_oam["CAP"], XBrushes.Black, false, null, 0, 380);
                                }
                            }
                            else
                            {
                                pdf.scrivi("Denominazione: " + dettaglio_oam["DENOMINAZIONE"], XBrushes.Black);
                                pdf.scrivi("Ndg: " + dettaglio_oam["NATURA GIURIDICA"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Codice Fiscale: " + dettaglio_oam["CODICE FISCALE"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Data Costituzione: " + dettaglio_oam["DATA COSTITUZIONE"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("PEC: " + dettaglio_oam["PEC"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Tipo Elenco: " + dettaglio_oam["TIPO ELENCO"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Numero Iscrizione: " + dettaglio_oam["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Stato: " + dettaglio_oam["STATO"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Autorizzato ad Operare: " + dettaglio_oam["AUTORIZZATO AD OPERARE"], XBrushes.Black, false, null, 0, 380);

                                pdf.scrivi("RAPPRESENTANTE LEGALE", XBrushes.Black, true, bold);
                                pdf.scrivi("Cognome e Nome: " + dettaglio_oam["COGNOME E NOME_RAPP_LEGALE"], XBrushes.Black);
                                pdf.scrivi("Sesso: " + dettaglio_oam["SESSO_RAPP_LEGALE"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Comune di Nascita: " + dettaglio_oam["COMUNE DI NASCITA_RAPP_LEGALE"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Codice Fiscale: " + dettaglio_oam["CODICE FISCALE_RAPP_LEGALE"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Provincia di Nascita: " + dettaglio_oam["PROVINCIA DI NASCITA_RAPP_LEGALE"], XBrushes.Black, false, null, 0, 380);
                                pdf.scrivi("Inizio Incarico: " + dettaglio_oam["INIZIO INCARICO_RAPP_LEGALE"], XBrushes.Black, false, null, -1);
                                pdf.scrivi("Data di Nascita: " + dettaglio_oam["DATA DI NASCITA_RAPP_LEGALE"], XBrushes.Black, false, null, 0, 380);

                                pdf.scrivi("SEDI", XBrushes.Black, true, bold);

                                var sedi = new List<(string Prefix, string Titolo, bool AddSpacing)>
                                {
                                    ("DIR_GENERALE", "DIREZIONE GENERALE", false),
                                    ("ITALIA", "SEDE ITALIANA", true),
                                    ("ESTERO", "SEDE ESTERA", true)
                                };

                                foreach ( var (prefix, titolo, spacing) in sedi)
                                {
                                    if (sedi_oam.ContainsKey($"COMUNE_{prefix}"))
                                    {
                                        SedeInfo sede = new SedeInfo
                                        {
                                            Titolo = titolo,
                                            Comune = sedi_oam[$"COMUNE_{prefix}"],
                                            Provincia = sedi_oam[$"PROVINCIA_{prefix}"],
                                            Indirizzo = sedi_oam[$"INDIRIZZO_{prefix}"],
                                            Cap = sedi_oam[$"CAP_{prefix}"]
                                        };
                                        ScriviSede(sede, spacing);
                                    }
                                }

                                /*
                                if (sedi_oam.ContainsKey("COMUNE_DIR_GENERALE"))
                                {
                                    pdf.scrivi("DIREZIONE GENERALE", XBrushes.Black, false, bold);

                                    pdf.scrivi("Comune: " + sedi_oam["COMUNE_DIR_GENERALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Provincia: " + sedi_oam["PROVINCIA_DIR_GENERALE"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Indirizzo: " + sedi_oam["INDIRIZZO_DIR_GENERALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("CAP: " + sedi_oam["CAP_DIR_GENERALE"], XBrushes.Black, false, null, 0, 380);
                                }

                                if (sedi_oam.ContainsKey("COMUNE_ITALIA"))
                                {
                                    pdf.addY(25);

                                    pdf.scrivi("SEDE ITALIANA", XBrushes.Black, false, bold);

                                    pdf.scrivi("Comune: " + sedi_oam["COMUNE_ITALIA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Provincia: " + sedi_oam["PROVINCIA_ITALIA"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Indirizzo: " + sedi_oam["INDIRIZZO_ITALIA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("CAP: " + sedi_oam["CAP_ITALIA"], XBrushes.Black, false, null, 0, 380);
                                }

                                if (sedi_oam.ContainsKey("COMUNE_ESTERO"))
                                {
                                    pdf.addY(25);

                                    pdf.scrivi("SEDE ESTERA", XBrushes.Black, false, bold);

                                    pdf.scrivi("Comune: " + sedi_oam["COMUNE_ITALIA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Provincia: " + sedi_oam["PROVINCIA_ITALIA"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Indirizzo: " + sedi_oam["INDIRIZZO_ITALIA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("CAP: " + sedi_oam["CAP_ITALIA"], XBrushes.Black, false, null, 0, 380);
                                }
                                */

                                bool linea = false;
                                pdf.entra(6);
                                pdf.scrivi("AMMINISTRAZIONE", XBrushes.Black, true, bold);
                                foreach (Dictionary<string, string> amministratore in amministratori_oam)
                                {
                                    if (linea)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    pdf.entra(5);
                                    pdf.scrivi(amministratore["NOMINATIVO"], XBrushes.Black, true, mini_bold);
                                    pdf.scrivi("Carica: " + amministratore["CARICA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Inizio Collaborazione: " + amministratore["INIZIO COLLABORAZIONE"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Luogo di Nascita: " + amministratore["LUOGO DI NASCITA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Codice Fiscale: " + amministratore["CODICE FISCALE"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Data di Nascita: " + amministratore["DATA DI NASCITA"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Sesso: " + amministratore["SESSO"], XBrushes.Black, false, null, 0, 380);

                                    linea = true;
                                }

                                pdf.entra(5);
                                pdf.scrivi("DIPENDENTI E COLLABORATORI", XBrushes.Black, true, bold);
                                linea = false;
                                foreach (Dictionary<string, string> dipendente in dipendenti_oam)
                                {
                                    if (linea)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    pdf.entra(4);

                                    pdf.scrivi(dipendente["NOMINATIVO"], XBrushes.Black, true, mini_bold);
                                    pdf.scrivi("Luogo di Nascita: " + dipendente["LUOGO DI NASCITA"], XBrushes.Black, false, null, -1);

                                    pdf.scrivi("Inizio Collaborazione: " + dipendente["INIZIO COLLABORAZIONE"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Data di Nascita: " + dipendente["DATA DI NASCITA"], XBrushes.Black, false, null, -1);

                                    pdf.scrivi("Numero Iscrizione: " + (dipendente.ContainsKey("NUMERO ISCRIZIONE") ? dipendente["NUMERO ISCRIZIONE"] : "NON PRESENTE"), XBrushes.Black, false, null, 0, 380);

                                    pdf.scrivi("Codice Fiscale: " + dipendente["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Sesso: " + dipendente["SESSO"], XBrushes.Black, false, null, 0, 380);

                                    linea = true;
                                }
                            }
                            indice_intermed = 0;

                            pdf.entra(3);
                            pdf.scrivi("MANDATI DIRETTI", XBrushes.Black, true, bold);
                            if (mandati_diretti_oam.Count > 0)
                            {
                                bool trovato_mandato = false;
                                foreach (Dictionary<string, string> mandato in mandati_diretti_oam)
                                {
                                    if (indice_intermed > 0)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    if (mandato["DENOMINAZIONE"] == "AVVERA S.P.A. (GIA' CREACASA)")
                                    {
                                        trovato_mandato = true;
                                    }

                                    pdf.entra(2);
                                    pdf.scrivi(mandato["DENOMINAZIONE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Appartiene a Gruppo: " + mandato["APPARTIENE A GRUPPO"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Codice Fiscale: " + mandato["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Inizio Mandato: " + mandato["INIZIO MANDATO"], XBrushes.Black, false, null, 0, 380);

                                    indice_prod = 0;
                                    prodotti = OAM.cattura_prodotti(indice_intermed, "Diretti");
                                    if (prodotti.Count > 0)
                                    {
                                        pdf.entra(3);
                                        pdf.scrivi("PRODOTTI E ATTIVITÀ", XBrushes.Black, true, bold);
                                        foreach (List<string> prod in prodotti)
                                        {
                                            pdf.scrivi("CODICE: " + prod[0] + " " + prod[1], XBrushes.Black);

                                            indice_prod += 1;
                                        }
                                    }
                                    indice_intermed += 1;
                                }

                                if (!trovato_mandato)
                                {
                                    riga_out.Add("");
                                }
                                else
                                {
                                    riga_out.Add("SI");
                                }
                            }
                            else
                            {
                                pdf.scrivi("Nessun Rapporto Attivo", XBrushes.Black, false, mini_bold);
                            }

                            indice_intermed = 0;
                            if (rapp_cessati_diretti.Count > 0)
                            {
                                pdf.entra(5);
                                pdf.scrivi("RAPPORTI CESSATI", XBrushes.Black, true, bold);
                                foreach (Dictionary<string, string> mandato in rapp_cessati_diretti)
                                {
                                    if (indice_intermed > 0)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    pdf.entra(3);
                                    pdf.scrivi(mandato["DENOMINAZIONE"], XBrushes.Black, false, mini_bold);
                                    pdf.scrivi("Codice Fiscale: " + mandato["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Inizio Mandato: " + mandato["INIZIO MANDATO"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Appartiene a Gruppo: " + mandato["APPARTIENE A GRUPPO"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Fine Mandato: " + (mandato.ContainsKey("FINE MANDATO") ? mandato["FINE MANDATO"] : ""), XBrushes.Black, false, null, 0, 380);

                                    indice_prod = 0;
                                    prodotti = new List<List<string>>();
                                    if (prodotti.Count > 0)
                                    {
                                        pdf.entra(4);
                                        pdf.scrivi("PRODOTTI E ATTIVITÀ", XBrushes.Black, true, bold);
                                        foreach (List<string> prod in prodotti)
                                        {
                                            pdf.scrivi("CODICE: " + prod[0] + " " + prod[1], XBrushes.Black);

                                            indice_prod += 1;
                                        }
                                    }

                                    indice_intermed += 1;
                                }
                            }
                            indice_intermed = 0;

                            pdf.entra(6);
                            pdf.scrivi("MANDATI INDIRETTI", XBrushes.Black, true, bold);
                            if (mandati_indiretti_oam.Count > 0)
                            {
                                foreach (Dictionary<string, string> mandato in mandati_indiretti_oam)
                                {
                                    pdf.entra(5);

                                    if (indice_intermed > 0)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    pdf.scrivi("AGENTE: " + mandato["DENOMINAZIONE"], XBrushes.Black, false, mini_bold);
                                    pdf.scrivi("Codice Fiscale: " + mandato["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Numero Iscrizione: " + mandato["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 0, 380);

                                    pdf.scrivi("INTERMEDIARIO: " + mandato["DENOMINAZIONE_INTERMED"], XBrushes.Black, false, mini_bold);
                                    pdf.scrivi("Codice Fiscale: " + mandato["CODICE FISCALE_INTERMED"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Inizio Mandato: " + mandato["INIZIO MANDATO"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Appartiene a Gruppo: " + mandato["APPARTIENE A GRUPPO"], XBrushes.Black);

                                    indice_prod = 0;
                                    prodotti = OAM.cattura_prodotti(indice_intermed, "Indiretti");
                                    if (prodotti.Count > 0)
                                    {
                                        pdf.entra(4);
                                        pdf.scrivi("PRODOTTI E ATTIVITÀ", XBrushes.Black, true, bold);
                                        foreach (List<string> prod in prodotti)
                                        {
                                            pdf.scrivi("CODICE: " + prod[0] + " " + prod[1], XBrushes.Black);

                                            indice_prod += 1;
                                        }
                                    }
                                    indice_intermed += 1;
                                }
                            }
                            else
                            {
                                pdf.scrivi("Nessun Rapporto Attivo", XBrushes.Black, false, mini_bold);
                            }

                            indice_intermed = 0;
                            if (rapp_cessati_indiretti.Count > 0)
                            {
                                pdf.entra(5);
                                pdf.scrivi("RAPPORTI CESSATI", XBrushes.Black, true, bold);
                                foreach (Dictionary<string, string> mandato in rapp_cessati_indiretti)
                                {
                                    if (indice_intermed > 0)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    pdf.entra(3);
                                    pdf.scrivi("AGENTE: " + mandato["DENOMINAZIONE"], XBrushes.Black, false, mini_bold);
                                    pdf.scrivi("Codice Fiscale: " + mandato["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Numero Iscrizione: " + mandato["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 0, 380);

                                    pdf.scrivi("INTERMEDIARIO: " + mandato["DENOMINAZIONE_INTERMED"], XBrushes.Black, false, mini_bold);
                                    pdf.scrivi("Codice Fiscale: " + mandato["CODICE FISCALE_INTERMED"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Inizio Mandato: " + mandato["INIZIO MANDATO"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Appartiene a Gruppo: " + mandato["APPARTIENE A GRUPPO"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Fine Mandato: " + mandato["FINE MANDATO"], XBrushes.Black, false, null, 0, 380);

                                    indice_prod = 0;
                                    prodotti = new List<List<string>>();
                                    if (prodotti.Count > 0)
                                    {
                                        pdf.entra(4);
                                        pdf.scrivi("PRODOTTI E ATTIVITÀ", XBrushes.Black, true, bold);
                                        foreach (List<string> prod in prodotti)
                                        {
                                            pdf.scrivi("CODICE: " + prod[0] + " " + prod[1], XBrushes.Black);

                                            indice_prod += 1;
                                        }
                                    }
                                    indice_intermed += 1;
                                }
                            }

                            if (storico_oam.Count > 0)
                            {
                                bool linea = false;
                                pdf.entra(4);
                                pdf.scrivi("STORICO STATI", XBrushes.Black, true, bold);
                                foreach (Dictionary<string, string> stato in storico_oam)
                                {
                                    if (linea)
                                    {
                                        pdf.entra(1);
                                        pdf.linea(1);
                                    }

                                    pdf.entra(3);
                                    pdf.scrivi("Data Stato: " + stato["DATA STATO"], XBrushes.Black);
                                    pdf.scrivi("Stato: " + stato["STATO"], XBrushes.Black, false, null, -1);
                                    pdf.scrivi("Numero Iscrizione: " + stato["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 0, 380);
                                    pdf.scrivi("Elenco: " + stato["ELENCO"], XBrushes.Black, false, null, -1);

                                    if (stato["AUTORIZZATO AD OPERARE"].Contains("NON AUTORIZZATO AD OPERARE"))
                                    {
                                        stato["AUTORIZZATO AD OPERARE"] = "NO";
                                    }
                                    pdf.scrivi("Autorizzato ad Operare: " + stato["AUTORIZZATO AD OPERARE"], XBrushes.Black, false, null, 0, 380);

                                    linea = true;
                                }
                            }
                        }
                        else
                        {
                            pdf.scrivi(dettaglio_oam["NOMINATIVO"], XBrushes.Black, true, bold);
                            pdf.scrivi("Codice Fiscale: " + dettaglio_oam["CODICE FISCALE"], XBrushes.Black, false, null, -1);
                            pdf.scrivi("Inizio Collaborazione: " + dettaglio_oam["INIZIO COLLABORAZIONE"], XBrushes.Black, false, null, 0, 380);

                            pdf.scrivi(dettaglio_oam["SOCIETA"], XBrushes.Black, true, bold);
                            pdf.scrivi("Tipo Elenco: " + dettaglio_oam["ELENCO"], XBrushes.Black, false, null, -1);
                            pdf.scrivi("Numero Iscrizione: " + dettaglio_oam["NUMERO ISCRIZIONE"], XBrushes.Black, false, null, 0, 380);

                            riga_out.Add("ISCRITTO");
                        }
                    }
                    else
                    {
                        pdf.scrivi(denominazione, XBrushes.Black, true, bold);
                        pdf.scrivi("NON ISCRITTO", XBrushes.Black, true, bold);

                        riga_out.Add("NON ISCRITTO");
                    }
                    pdf.salva(Utils.CurDir() + "\\PDF\\", denominazione + "_" + codice_fiscale + "_OAM.pdf");
                #endregion

                prossimo:
                    excel.WriteExcel(riga_out, riga, 5);
                    riga_out = new List<string>();
                    riga += 1;
                }
                excel.Close();
                browser.Quit();
                MsgBox.Show("Elaborazione Terminata!");
            }
            catch (Exception ex)
            {
                excel.Close();
                browser?.Close();
                log.Write(ex.Message, 4, ex);
                win.Chiudi();
            }
        }
        private static string RimpiazzaAccenti(string denominazione, string lettera, int ascii)
        {
            return denominazione.Replace(Encoding.Default.GetString(new byte[] { (byte)ascii }), lettera + "'");
        }
        private static void ScriviSede(SedeInfo sede, bool addExtraSpacing = false)
        {
            if (addExtraSpacing)
                pdf.addY(25);

            pdf.scrivi(sede.Titolo, XBrushes.Black, false, bold);
            pdf.scrivi("Comune: " + sede.Comune, XBrushes.Black, false, null, -1);
            pdf.scrivi("Provincia: " + sede.Provincia, XBrushes.Black, false, null, 0, 380);
            pdf.scrivi("Indirizzo: " + sede.Indirizzo, XBrushes.Black, false, null, -1);
            pdf.scrivi("CAP: " + sede.Cap, XBrushes.Black, false, null, 0, 380);
        }
    }
    public class SedeInfo
    {
        public string Titolo { get; set; }
        public string Comune { get; set; }
        public string Provincia { get; set; }
        public string Indirizzo { get; set; }
        public string Cap { get; set; }
    }
}

