using DLL;
using Microsoft.Win32;
using OpenQA.Selenium.Chrome;

namespace Deutsche
{
    internal class Deutsche
    {
        [STAThread]
        static void Main()
        {
            Win win = new Win("Deutsche");
            ChromeDriver browser = null;

            List<List<string>> riga_out = new List<List<string>>();
            List<string> scrivi = new List<string>();

            LogFile log = new LogFile();
            LogFile log_cli = new LogFile("LOG_Cliente" + DateTime.Now.ToString("_yyyyMMdd_HHmm") + ".log");

            Excel excel = new Excel();
            Excel excel_out = new Excel();

            try
            {
                OpenFileDialog scegli_file = new OpenFileDialog();
                scegli_file.Filter = "Foglio di lavoro Excel (*.xls*) | *.xls*";
                scegli_file.Title = "Seleziona File Elenco Agenti per Crif";

                if (scegli_file.ShowDialog() == true)
                {
                    int riga = 0;
                    try
                    {
                        InputBox inp_box = new InputBox("Riga inizio");
                        riga = inp_box.result != "" ? int.Parse(inp_box.result) : 2;
                    }
                    catch (Exception)
                    {
                        Environment.Exit(0);
                    }
                    string[] riga_letta = new string[] { };

                    string elencoAgentiCrif = scegli_file.FileName;
                    excel.Open(elencoAgentiCrif);
                    excel_out.Open(Utils.CurDir() + "\\Deutsche_Output.xlsx");
                    MessageFrame msg = new MessageFrame(1);

                    browser = HTML.ApriBrowser_CR("https://www.organismo-am.it/elenchi-registri/filtri.html");
                    OAM.collega_driver(browser);
                    while (true)
                    {
                        browser.Navigate().GoToUrl("https://www.organismo-am.it/elenchi-registri/filtri.html");
                        msg.scrivi("Elaboro riga " + riga);

                        riga_letta = excel.ReadExcel(riga);

                        log.Write("Lavoro riga " + riga + " soggetto: " + riga_letta[1]);

                        string denominazione = riga_letta[1];
                        if (denominazione == "")
                        {
                            break;
                        }
                        string comune = riga_letta[4];
                        string cod_fis = riga_letta[6];
                        string niscr = riga_letta[7];
                        string nome_foglio = riga_letta[8];
                        string lavorata = riga_letta[9];

                        if (lavorata != "")
                        {
                            riga += 1;
                            continue;
                        }

                        //goto ivass;

                        log_cli.Write("===================================================================================================");
                        log_cli.Write("Inizio lavorazione soggetto " + denominazione);
                        log_cli.Write("Accesso ad OAM alle ore " + DateTime.Now.ToString("HH:mm:ss") + " del " + DateTime.Now.ToString("dd/MM/yyyy"));

                        if (OAM.ricerca(cod_fis))
                        {
                            Thread.Sleep(1000);
                            //List<string> dettaglio_oam = new List<string>();
                            Dictionary<string, string> dettaglio_oam = new Dictionary<string, string>();

                            while (true)
                            {
                                dettaglio_oam = OAM.cattura_dettaglio();
                                if (dettaglio_oam.Count == 0)
                                {
                                    Thread.Sleep(1000);
                                    continue;
                                }
                                else
                                {
                                    break;
                                }
                            }
                            //List<string> sedi = oam.cattura_sedi();
                            //List<List<string>> storico = oam.cattura_storico();
                            //List<List<string>> mandati = oam.cattura_mandati("Diretti");

                            Dictionary<string, string> sedi = OAM.cattura_sedi();
                            List<Dictionary<string, string>> storico = OAM.cattura_storico();
                            List<Dictionary<string, string>> mandati = OAM.cattura_mandati("Diretti");

                            string indirizzo = string.Empty;
                            if (sedi.Count > 0)
                            {
                                indirizzo = string.Format("{0} - {1} {2} ({3})", sedi["INDIRIZZO_DIR_GENERALE"], sedi["CAP_DIR_GENERALE"], sedi["COMUNE_DIR_GENERALE"], sedi["PROVINCIA_DIR_GENERALE"]);
                            }
                            else
                            {
                                indirizzo = string.Format("{0} - {1} {2} ({3})", dettaglio_oam["INDIRIZZO"], dettaglio_oam["CAP"], dettaglio_oam["COMUNE"], dettaglio_oam["PROVINCIA"]);
                            }

                            scrivi.Add(dettaglio_oam.ContainsKey("DENOMINAZIONE") ?
                                        dettaglio_oam["DENOMINAZIONE"] : dettaglio_oam["COGNOME E NOME"]);
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "" };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "ISCRIZIONE OAM" };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { (cod_fis.Length > 11) ? "Cognome e Nome" : "Denominazione Sociale",
                                        dettaglio_oam.ContainsKey("DENOMINAZIONE") ?
                                        dettaglio_oam["DENOMINAZIONE"] : dettaglio_oam["COGNOME E NOME"] };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Persona", (cod_fis.Length > 11) ? "Fisica" : "Giuridica" };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Codice fiscale", "'" + dettaglio_oam["CODICE FISCALE"] };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Domicilio / Sede legale", indirizzo };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Elenco", dettaglio_oam["TIPO ELENCO"] };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "N° Iscrizione", dettaglio_oam["NUMERO ISCRIZIONE"] };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Stato", dettaglio_oam["STATO"] };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Data delibera", "'" + storico[0]["DATA STATO"] };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "PersonType" };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "Type" };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "" };
                            riga_out.Add(scrivi);
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "INFORMAZIONI PRINCIPALI" };
                            riga_out.Add(scrivi);
                            scrivi = new List<string>() { "DATI ANAGRAFICI" };
                            riga_out.Add(scrivi);


                            switch (cod_fis)
                            {
                                case string x when x.Length > 11:
                                    scrivi = new List<string>() { "Denominazione Persona Fisica", dettaglio_oam["COGNOME E NOME"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Sesso", dettaglio_oam["SESSO"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Codice Fiscale", dettaglio_oam["CODICE FISCALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Data di nascita", "'" + dettaglio_oam["DATA DI NASCITA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Comune di nascita", dettaglio_oam["COMUNE DI NASCITA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Provincia di nascita", dettaglio_oam["PROVINCIA DI NASCITA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Cittadinanza", dettaglio_oam["CITTADINANZA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Denominazione Ditta", dettaglio_oam.ContainsKey("DITTA") ? dettaglio_oam["DITTA"] : "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "PEC", dettaglio_oam["PEC"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "RESIDENZA" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Indirizzo", dettaglio_oam["INDIRIZZO"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "CAP", dettaglio_oam["CAP"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Comune", dettaglio_oam["COMUNE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Provincia", dettaglio_oam["PROVINCIA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "DOMICILIO" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Indirizzo", dettaglio_oam["INDIRIZZO"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "CAP", dettaglio_oam["CAP"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Comune", dettaglio_oam["COMUNE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Provincia", dettaglio_oam["PROVINCIA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);
                                    riga_out.Add(scrivi);

                                    break;

                                case string x when x.Length == 11:
                                    scrivi = new List<string>() { "Denominazione Sociale", dettaglio_oam["DENOMINAZIONE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Natura Giuridica", dettaglio_oam["NATURA GIURIDICA"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Codice Fiscale", "'" + dettaglio_oam["CODICE FISCALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Data Costituzione", "'" + dettaglio_oam["DATA COSTITUZIONE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "PEC", dettaglio_oam["PEC"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "RAPPRESENTANTE LEGALE" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Cognome e Nome", dettaglio_oam["COGNOME E NOME_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Sesso", dettaglio_oam["SESSO_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Codice Fiscale", dettaglio_oam["CODICE FISCALE_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Provincia di Nascita", dettaglio_oam["PROVINCIA DI NASCITA_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Comune di Nascita", dettaglio_oam["COMUNE DI NASCITA_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Data di Nascita", "'" + dettaglio_oam["DATA DI NASCITA_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Inizio Incarico", "'" + dettaglio_oam["INIZIO INCARICO_RAPP_LEGALE"] };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "SEDI" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "DIREZIONE GENERALE", "", "", "SEDE ITALIANA", "", "", "SEDE ESTERA" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Indirizzo", sedi["INDIRIZZO_DIR_GENERALE"], "",
                                                              "Indirizzo", sedi.ContainsKey("INDIRIZZO_ITALIA") ? sedi["INDIRIZZO_ITALIA"] : "", "",
                                                              "Indirizzo", sedi.ContainsKey("INDIRIZZO_ESTERO") ? sedi["INDIRIZZO_ESTERO"] : "", "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "CAP", sedi["CAP_DIR_GENERALE"], "",
                                                              "CAP", sedi.ContainsKey("CAP_ITALIA") ? sedi["CAP_ITALIA"] : "", "",
                                                              "CAP", sedi.ContainsKey("CAP_ESTERO") ? sedi["CAP_ESTERO"] : "", "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Comune", sedi["COMUNE_DIR_GENERALE"], "",
                                                              "Comune", sedi.ContainsKey("COMUNE_ITALIA") ? sedi["COMUNE_ITALIA"] : "", "",
                                                              "Comune", sedi.ContainsKey("COMUNE_ESTERO") ? sedi["COMUNE_ESTERO"] : "", "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "Provincia", sedi["PROVINCIA_DIR_GENERALE"], "",
                                                              "Provincia", sedi.ContainsKey("PROVINCIA_ITALIA") ? sedi["PROVINCIA_ITALIA"] : "", "",
                                                              "Provincia", sedi.ContainsKey("PROVICNIA_ESTERO") ? sedi["PROVINCIA_ESTERO"] : "", "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "", "", "", "Telefono", "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "", "", "", "Fax", "" };
                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);

                                    break;
                            }

                            int indice_intermed = 0;
                            foreach (Dictionary<string, string> intermediario in mandati)
                            {
                                List<List<string>> prodotti = OAM.cattura_prodotti(indice_intermed);
                                scrivi = new List<string>() { "INTERMEDIARI PREPONENTI", "MANDATI DIRETTI" };
                                riga_out.Add(scrivi);
                                scrivi = new List<string>() { "Denominazione", intermediario["DENOMINAZIONE"] };
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                for (int i = 0; i < 3; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            scrivi.Add("Tipologia");
                                            break;

                                        case 1:
                                            scrivi.Add("Codice");
                                            break;

                                        case 2:
                                            scrivi.Add("Descrizione");
                                            break;
                                    }

                                    foreach (List<string> prod in prodotti)
                                    {
                                        if (i == 0)
                                        {
                                            scrivi.Add("PRODOTTI");
                                        }
                                        else
                                        {
                                            scrivi.Add(prod[i - 1]);
                                        }
                                    }

                                    riga_out.Add(scrivi);
                                    scrivi = new List<string>();
                                }

                                scrivi = new List<string>() { "Codice Fiscale", "'" + intermediario["CODICE FISCALE"] };
                                riga_out.Add(scrivi);
                                scrivi = new List<string>() { "Appartiene a Gruppo", intermediario["APPARTIENE A GRUPPO"] };
                                riga_out.Add(scrivi);
                                scrivi = new List<string>() { "Inizio Collaborazione", "'" + intermediario["INIZIO MANDATO"] };
                                riga_out.Add(scrivi);
                                scrivi = new List<string>() { "" };
                                riga_out.Add(scrivi);

                                indice_intermed += 1;
                            }

                            //List<List<string>> amministratori = oam.cattura_amministratori();
                            List<Dictionary<string, string>> amministratori = OAM.cattura_amministratori();
                            if (amministratori.Count > 0)
                            {
                                riga_out.Add(scrivi);
                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("AMMINISTRAZIONE");
                                    scrivi.Add("");
                                    scrivi.Add("");

                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Cognome e Nome");
                                    scrivi.Add(amm["NOMINATIVO"]);
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Luogo di nascita");
                                    scrivi.Add(amm["LUOGO DI NASCITA"]);
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Data di Nascita");
                                    scrivi.Add("'" + amm["DATA DI NASCITA"]);
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Sesso");
                                    scrivi.Add(amm["SESSO"]);
                                    scrivi.Add("");

                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Codice Fiscale");
                                    scrivi.Add(amm["CODICE FISCALE"]);
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Carica");
                                    scrivi.Add(amm["CARICA"]);
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Denominazione Ditta");
                                    scrivi.Add("");
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Partita Iva");
                                    scrivi.Add("");
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);
                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Ruolo");
                                    scrivi.Add("");
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>();
                                foreach (Dictionary<string, string> amm in amministratori)
                                {
                                    scrivi.Add("Inizio Collaborazione");
                                    scrivi.Add("'" + amm["INIZIO COLLABORAZIONE"]);
                                    scrivi.Add("");
                                }
                                riga_out.Add(scrivi);

                                scrivi = new List<string>() { "" };
                                riga_out.Add(scrivi);
                            }

                            //List<List<string>> dipendenti = oam.cattura_dipendenti();
                            List<Dictionary<string, string>> dipendenti = OAM.cattura_dipendenti();
                            //log.Write("Dopo cattura dipendenti alle ore " + DateTime.Now.ToString("HH:mm:ss") + " del " + DateTime.Now.ToString("dd/MM/yyyy") + " n." + dipendenti.Count);
                            if (dipendenti.Count > 0)
                            {
                                scrivi = new List<string>() { "DIPENDENTI E COLLABORATORI" };
                                riga_out.Add(scrivi);

                                scrivi = new List<string>() { "Cognome e Nome", "Luogo di nascita", "Data di Nascita", "Sesso", "Codice Fiscale", "Numero Iscrizione", "Note", "Inizio Collaborazione", "Amministratore" };
                                riga_out.Add(scrivi);

                                foreach (Dictionary<string, string> dipen in dipendenti)
                                {
                                    scrivi = new List<string>() { dipen["NOMINATIVO"], dipen["LUOGO DI NASCITA"], "'" + dipen["DATA DI NASCITA"], dipen["SESSO"], dipen["CODICE FISCALE"],
                                                            dipen.ContainsKey("NUMERO ISCRIZIONE") ? dipen["NUMERO ISCRIZIONE"] : "",
                                                            dipen.ContainsKey("NOTE") ? dipen["NOTE"] : "",
                                                            "'" + dipen["INIZIO COLLABORAZIONE"], dipen.ContainsKey("AMMINISTRATORE") ? dipen["AMMINISTRATORE"] : "" };
                                    riga_out.Add(scrivi);
                                }

                                scrivi = new List<string>() { "" };
                                riga_out.Add(scrivi);
                            }
                        }
                        else
                        {

                        }
                    ivass:
                        browser.Navigate().GoToUrl("https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                        Ivass.collega_driver(browser);
                        scrivi = new List<string>() { "ISCRIZIONE IVASS/RUI" };
                        riga_out.Add(scrivi);

                        Dictionary<string, string> dettaglio_ivass = Ivass.ricerca(niscr.Trim());
                        log_cli.Write("Accesso ad IVASS alle ore " + DateTime.Now.ToString("HH:mm:ss") + " del " + DateTime.Now.ToString("dd/MM/yyyy"));

                        string tmp = string.Empty;
                        if (dettaglio_ivass.Count > 0)
                        {
                            switch (cod_fis)
                            {
                                case string x when x.Length > 11:
                                    scrivi = new List<string>() { "Numero Iscrizione", dettaglio_ivass["NUMERO ISCRIZIONE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Sezione", dettaglio_ivass["SEZIONE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Nominativo", dettaglio_ivass["NOMINATIVO"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Data Iscrizione", dettaglio_ivass["DATA ISCRIZIONE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Luogo Nascita", dettaglio_ivass["LUOGO NASCITA"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Data di Nascita", dettaglio_ivass["DATA NASCITA"] };
                                    riga_out.Add(scrivi);

                                    tmp = string.Empty;
                                    if (dettaglio_ivass.ContainsKey("QUALIFICA DI ESERCIZIO"))
                                    {
                                        tmp = dettaglio_ivass["QUALIFICA DI ESERCIZIO"];
                                    }
                                    scrivi = new List<string>() { "QUALIFICA DI ESERCIZIO", tmp };
                                    riga_out.Add(scrivi);

                                    tmp = string.Empty;
                                    if (dettaglio_ivass.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                                    {
                                        tmp = dettaglio_ivass["INTERMEDIARI PER CUI OPERA"];
                                    }
                                    scrivi = new List<string>() { "INTERMEDIARI PER CUI OPERA", tmp };
                                    riga_out.Add(scrivi);

                                    break;

                                case string x when x.Length == 11:
                                    scrivi = new List<string>() { "Numero Iscrizione", dettaglio_ivass["NUMERO ISCRIZIONE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Sezione", dettaglio_ivass["SEZIONE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Ragione o denominazione sociale", dettaglio_ivass["RAGIONE O DENOMINAZIONE SOCIALE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Data Iscrizione", dettaglio_ivass["DATA ISCRIZIONE"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "Sede Legale", dettaglio_ivass["SEDE LEGALE"].Trim() };
                                    riga_out.Add(scrivi);

                                    tmp = "";
                                    if (dettaglio_ivass.ContainsKey("RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE"))
                                    {
                                        tmp = dettaglio_ivass["RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE"];
                                    }
                                    scrivi = new List<string>() { "RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE", tmp };
                                    riga_out.Add(scrivi);

                                    tmp = "";
                                    if (dettaglio_ivass.ContainsKey("ADDETTI ALL'ATTIVITÀ DI INTERMEDIAZIONE"))
                                    {
                                        tmp = dettaglio_ivass["ADDETTI ALL'ATTIVITÀ DI INTERMEDIAZIONE"];
                                    }
                                    scrivi = new List<string>() { "ADDETTI ALL'ATTIVITÀ DI INTERMEDIAZIONE", tmp };
                                    riga_out.Add(scrivi);

                                    tmp = "";
                                    if (dettaglio_ivass.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                                    {
                                        tmp = dettaglio_ivass["INTERMEDIARI PER CUI OPERA"];
                                    }
                                    scrivi = new List<string>() { "INTERMEDIARI PER CUI OPERA", dettaglio_ivass["INTERMEDIARI PER CUI OPERA"] };
                                    riga_out.Add(scrivi);

                                    scrivi = new List<string>() { "" };
                                    riga_out.Add(scrivi);


                                    List<List<string>> temp = new List<List<string>>();
                                    List<List<List<string>>> temp_2 = new List<List<List<string>>>();

                                    if (dettaglio_ivass.ContainsKey("RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE"))
                                    {
                                        foreach (string resp in dettaglio_ivass["RESPONSABILI DELL'ATTIVITÀ DI INTERMEDIAZIONE"].Split('\n'))
                                        {
                                            string cerca = Utils.ExtractBetween(resp, "Numero iscrizione", "").Trim();
                                            dettaglio_ivass = Ivass.ricerca(cerca);
                                            if (dettaglio_ivass.Count > 0)
                                            {
                                                scrivi = new List<string>() { "ISCRIZIONE IVASS/RUI", "" };
                                                temp.Add(scrivi);

                                                scrivi = new List<string>() { "Numero Iscrizione", dettaglio_ivass["NUMERO ISCRIZIONE"] };
                                                temp.Add(scrivi);

                                                scrivi = new List<string>() { "Sezione", dettaglio_ivass["SEZIONE"] };
                                                temp.Add(scrivi);

                                                scrivi = new List<string>() { "Nominativo", dettaglio_ivass["NOMINATIVO"] };
                                                temp.Add(scrivi);

                                                scrivi = new List<string>() { "Data Iscrizione", "'" + dettaglio_ivass["DATA ISCRIZIONE"] };
                                                temp.Add(scrivi);

                                                tmp = "";
                                                if (dettaglio_ivass.ContainsKey("LUOGO NASCITA"))
                                                {
                                                    tmp = dettaglio_ivass["LUOGO NASCITA"];
                                                }
                                                scrivi = new List<string>() { "LUOGO DI NASCITA", tmp };
                                                temp.Add(scrivi);

                                                tmp = "";
                                                if (dettaglio_ivass.ContainsKey("DATA NASCITA"))
                                                {
                                                    tmp = dettaglio_ivass["DATA NASCITA"];
                                                }
                                                scrivi = new List<string>() { "Data di Nascita", "'" + dettaglio_ivass["DATA NASCITA"] };
                                                temp.Add(scrivi);

                                                tmp = "";
                                                if (dettaglio_ivass.ContainsKey("QUALIFICA DI ESERCIZIO"))
                                                {
                                                    tmp = dettaglio_ivass["QUALIFICA DI ESERCIZIO"];
                                                }
                                                scrivi = new List<string>() { "QUALIFICA DI ESERCIZIO", tmp };
                                                temp.Add(scrivi);

                                                tmp = "";
                                                if (dettaglio_ivass.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                                                {
                                                    tmp = dettaglio_ivass["INTERMEDIARI PER CUI OPERA"];
                                                }
                                                scrivi = new List<string>() { "INTERMEDIARI PER CUI OPERA", tmp };
                                                temp.Add(scrivi);

                                                temp_2.Add(temp);
                                                temp = new List<List<string>>();
                                            }
                                        }

                                        int z = 0;
                                        while (z < 9)
                                        {
                                            scrivi = new List<string>();
                                            foreach (List<List<string>> resp_intermed in temp_2)
                                            {
                                                foreach (string str in resp_intermed[z])
                                                {
                                                    scrivi.Add(str);
                                                }
                                                scrivi.Add("");
                                            }
                                            riga_out.Add(scrivi);
                                            z += 1;
                                        }
                                    }
                                    break;
                            }
                        }
                        else
                        {
                            scrivi = new List<string>() { "Soggetto", "NON ISCRITTO IN RUI" };
                            riga_out.Add(scrivi);
                        }

                        int riga_output = 1;
                        foreach (List<string> output in riga_out)
                        {
                            excel_out.WriteExcel(output, riga_output, 1, riga - 1);
                            riga_output += 1;
                        }
                        excel.WriteExcel(["1"], riga, 10);

                        scrivi = new List<string>();
                        riga_out = new List<List<string>>();

                        riga += 1;
                        //riga = 63;
                        log_cli.Write("Fine elaborazione soggetto");
                    }

                    excel.Close();
                    excel_out.Close();
                    browser.Close();
                }
                MsgBox.Show("Elaborazione Terminata!");
            }
            catch (Exception ex)
            {
                HTML.chiudiDriver();
                excel.Close();
                excel_out.Close();
                log.Write(ex.Message, 4, ex);
                win.Chiudi();
            }
        }
    }
}


