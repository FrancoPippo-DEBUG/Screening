using DLL;
using OpenQA.Selenium.Chrome;

namespace CompassOAM
{
    internal class CompassOAM
    {
        [STAThread]
        static void Main()
        {
            Win win = new Win("Compass OAM");
            ChromeDriver browser = null;

            var log = new LogFile();
            var excel = new Excel();

            try
            {
                int foglio_excel = 1;

                InputBox input = new("Riga partenza", "Compass OAM");
                if (!input.ret)
                {
                    return;
                }

                int riga = int.Parse(input.result == "" ? "2" : input.result);
                string[] riga_letta = new string[] { };
                List<List<string>> result = new List<List<string>>();
                List<string> riga_out = new List<string>();

                Dictionary<string, string> dettagli_oam = new Dictionary<string, string>();
                List<Dictionary<string, string>> mandati = new List<Dictionary<string, string>>();
                List<Dictionary<string, string>> dipendenti = new List<Dictionary<string, string>>();

                excel.Open(Utils.CurDir() + "\\OAM_Compass_Lavoro.xlsx");
                MessageFrame msg = new MessageFrame(1);

                string nome = string.Empty;
                string cognome = string.Empty;
                string codice_fiscale = string.Empty;
                string data_nascita = string.Empty;
                string luogo_nascita = string.Empty;
                string soc_coll = string.Empty;
                string piva_soc = string.Empty;
                string piva_soc_old = string.Empty;
                string intermediario = string.Empty;
                string piva_intermed = string.Empty;
                string piva_intermed_old = string.Empty;

                string tipo_intermed = string.Empty;
                bool mediatore = false;
                bool senza_cf = true;

                List<string> intermed_prec = new List<string>();
                browser = HTML.ApriBrowser_CR("https://www.google.it");

                while (foglio_excel > 0)
                {
                    while (true)
                    {
                        log.Write("----------- Elaboro riga " + riga);
                        msg.scrivi("Elaboro riga " + riga);

                        riga_letta = excel.ReadExcel(riga, foglio_excel);


                        if (riga_letta[13] != "")
                        {
                            riga += 1;
                            continue;
                        }

                        if (riga_letta[1] == "")
                        {
                            break;
                        }

                        nome = riga_letta[1];
                        cognome = riga_letta[2];
                        if (cognome.Contains("' "))
                        {
                            cognome = cognome.Replace("' ", "'");
                        }
                        codice_fiscale = riga_letta[3].Trim();
                        data_nascita = DateTime.Parse(riga_letta[4]).ToString("dd/MM/yyyy");
                        luogo_nascita = riga_letta[5];
                        soc_coll = riga_letta[7];

                        piva_soc = riga_letta[8];
                        intermediario = riga_letta[10];
                        piva_intermed = riga_letta[11];

                        tipo_intermed = riga_letta[12];

                        riga_out.Add(DateTime.Now.ToString("dd/MM/yyyy"));

                        string nome_ocf = nome;
                        switch (tipo_intermed)
                        {
                            case "BANCA+ PROMOTORI": // FOGLIO PROMOTORI
                                browser.Navigate().GoToUrl("https://www.organismocf.it/portal/web/portale-ocf/ricerca-nelle-sezioni-dell-albo");
                                OCF.collega_driver(browser);

                                while (true)
                                {
                                    Dictionary<string, string> dettagli_ocf = OCF.cerca(cognome, nome_ocf, data_nascita);
                                    if (dettagli_ocf.Count > 0)
                                    {
                                        riga_out.Add("1");
                                        riga_out.Add("");
                                        riga_out.Add("");
                                        riga_out.Add("ISCRITTO APF");
                                        riga_out.Add("");
                                        riga_out.Add("");
                                        riga_out.Add("NON ISCRITTO");
                                        riga_out.Add("");
                                        riga_out.Add("0");
                                        riga_out.Add("");
                                        riga_out.Add("");
                                        riga_out.Add("NON ISCRITTO");

                                        List<string> lista_out = new List<string>();
                                        bool trovato_intermed = false;
                                        Thread.Sleep(2000);
                                        List<Dictionary<string, string>> intermediari_ocf = OCF.intermediari();
                                        if (intermediari_ocf.Count > 0)
                                        {
                                            //for (int i = 0; i < 3; i++)
                                            foreach (Dictionary<string, string> intermed in intermediari_ocf)
                                            {
                                                if (pulisci_ragsoc(intermed["SOGGETTO"]) ==
                                                    pulisci_ragsoc(intermediario))
                                                {
                                                    riga_out.Add(intermed["SOGGETTO"]);
                                                    riga_out.Add("1");
                                                    riga_out.Add("'" + intermed["DATA INIZIO"]);
                                                    riga_out.Add("");
                                                    riga_out.Add("");

                                                    trovato_intermed = true;
                                                    for (int indice_prod = 0; indice_prod < 10; indice_prod++)
                                                    {
                                                        riga_out.Add("");
                                                    }
                                                }
                                                else
                                                {
                                                    lista_out.Add(intermed["SOGGETTO"]);
                                                    lista_out.Add("'" + intermed["DATA INIZIO"]);
                                                    lista_out.Add("");
                                                    lista_out.Add("");
                                                    for (int indice_prod = 0; indice_prod < 10; indice_prod++)
                                                    {
                                                        lista_out.Add("");
                                                    }
                                                }
                                            }

                                            if (lista_out.Count == 0)
                                            {
                                                for (int i = 0; i < 28; i++)
                                                {
                                                    lista_out.Add("");
                                                }
                                            }

                                            //foreach (string x in lista_out)
                                            int indice_out = 0;
                                            int x = 0;
                                            while (riga_out.Count < 56)
                                            {
                                                if (lista_out.Count > x)
                                                {
                                                    if (!trovato_intermed && indice_out == 1)
                                                    {
                                                        riga_out.Add("0");
                                                    }
                                                    riga_out.Add(lista_out[x]);

                                                    indice_out += 1;
                                                    if (trovato_intermed && indice_out >= 28)
                                                    {
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    riga_out.Add("");
                                                }

                                                x += 1;
                                            }
                                        }
                                        else
                                        {
                                            riga_out.Add("");
                                            riga_out.Add("0");
                                            for (int i = 0; i < 56; i++)
                                            {
                                                riga_out.Add("");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (nome_ocf.Contains(" "))
                                        {
                                            nome_ocf = Utils.sina(nome_ocf, " ");
                                            continue;
                                        }
                                        else
                                        {
                                            riga_out.Add("0");
                                            riga_out.Add("");
                                            riga_out.Add("");
                                            riga_out.Add("NON ISCRITTO APF");
                                            riga_out.Add("");
                                            riga_out.Add("");
                                            riga_out.Add("NON ISCRITTO");
                                            riga_out.Add("");
                                            riga_out.Add("0");
                                            riga_out.Add("");
                                            riga_out.Add("");
                                            riga_out.Add("NON ISCRITTO");
                                            riga_out.Add("");
                                            riga_out.Add("0");
                                            for (int i = 0; i < 41; i++)
                                            {
                                                riga_out.Add("");
                                            }
                                        }
                                    }
                                    break;
                                }
                                break;

                            case "MEDIATORE": // FOGLIO MEDIATORI
                                if (piva_intermed_old != piva_intermed) // SE PIVA INTERMED != PIVA PREC
                                { // CATTURO DETTAGLI INTERMEDIARIO
                                    browser.Navigate().GoToUrl("https://www.organismo-am.it/elenchi-registri/filtri.html");
                                    OAM.collega_driver(browser);

                                    if (OAM.ricerca(piva_intermed))
                                    {
                                        while (true)
                                        {
                                            dettagli_oam = OAM.cattura_dettaglio();
                                            if (dettagli_oam.Count > 0)
                                            {
                                                break;
                                            }
                                        }

                                        dipendenti = OAM.cattura_dipendenti();
                                    }
                                }

                                bool alert_coll = false;
                                foreach (Dictionary<string, string> dipen in dipendenti)
                                {
                                    if (dipen["CODICE FISCALE"] == codice_fiscale.ToUpper())
                                    {
                                        riga_out.Add("1");                                // COLONNA O - ALERT COLLABORATORE     
                                        riga_out.Add(dipen["INIZIO COLLABORAZIONE"]);     // COLONNA P - DATA INIZIO COLLABO
                                        riga_out.Add(dipen["CODICE FISCALE"]);            // COLONNA Q - CF COLLABORATORE
                                        riga_out.Add("");                                 // COLONNA R - NOTE

                                        riga_out.Add(dipen.ContainsKey("NUMERO ISCRIZIONE") ?
                                                        dipen["NUMERO ISCRIZIONE"] : ""); // COLONNA S - N ISCRIZ COLLABORATORE

                                        riga_out.Add("");           // COLONNA T - DATA ISCR DIR COLLABO
                                        riga_out.Add("");           // COLONNA U - STATO OAM COLL

                                        riga_out.Add("");           // COLONNA V - DENOM SOC COLL
                                        riga_out.Add("");           // COLONNA W - ALERT SOC COLL
                                        riga_out.Add("");           // COLONNA X - N ISCR SOC COLL
                                        riga_out.Add("");           // COLONNA Y - TIPO SOC COLL
                                        riga_out.Add("");           // COLONNA Z - STATO SOC COLL

                                        riga_out.Add(dettagli_oam["DENOMINAZIONE"]);  // COLONNA AA - DENOM INTERMED                                        
                                        riga_out.Add("1");              // COLONNA AB - ALERT INTERMED                                        
                                        riga_out.Add(dipen["INIZIO COLLABORAZIONE"]);         // COLONNA AC - DATA INIZIO INTERMED
                                        riga_out.Add(dettagli_oam["NUMERO ISCRIZIONE"]);  // COLONNA AD - N ISCR INTERMED
                                        riga_out.Add(dettagli_oam["TIPO ELENCO"]);  // COLONNA AE - SEZIONE INTERMED

                                        for (int i = 0; i < 38; i++)
                                        {
                                            riga_out.Add("");
                                        }

                                        alert_coll = true;
                                        break;
                                    }
                                }
                                
                                if (!alert_coll)
                                {
                                    riga_out.Add("0");              // COLONNA O - ALERT COLLABORATORE     
                                    for (int i = 0; i < 54; i++)
                                    {
                                        riga_out.Add("");
                                    }
                                }

                                break;

                            default: // FOGLIO 106 & 107
                                browser.Navigate().GoToUrl("https://www.organismo-am.it/elenchi-registri/filtri.html");
                                OAM.collega_driver(browser);

                                string cf_ricerca = piva_soc;
                                /*
                                if (tipo_intermed == "AGENTE")
                                {
                                    cf_ricerca = piva_intermed;
                                }
                                */
                                if (cf_ricerca.Trim() == "")
                                {
                                    senza_cf = true;
                                    continue;
                                }

                                bool trovato = false;
                                while (true)
                                {
                                    trovato = OAM.ricerca(cf_ricerca);
                                    if (!trovato && tipo_intermed == "AGENTE" && cf_ricerca != piva_intermed)
                                    {
                                        cf_ricerca = piva_intermed;
                                        continue;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                if (trovato)
                                {
                                    while (true)
                                    {
                                        dettagli_oam = OAM.cattura_dettaglio();
                                        if (dettagli_oam.Count == 0 && !dettagli_oam.ContainsKey("STATO"))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }

                                    if (piva_soc != piva_soc_old)
                                    {
                                        dipendenti = OAM.cattura_dipendenti();
                                        mandati = OAM.cattura_mandati("Diretti");
                                    }

                                    alert_coll = false;

                                    if (dettagli_oam["CODICE FISCALE"] == codice_fiscale.ToUpper())
                                    {
                                        riga_out.Add("1");
                                        riga_out.Add("");
                                        riga_out.Add(dettagli_oam["CODICE FISCALE"]);
                                        riga_out.Add("");

                                        riga_out.Add(dettagli_oam["NUMERO ISCRIZIONE"]);
                                        riga_out.Add("");
                                        riga_out.Add(dettagli_oam["STATO"]);

                                        alert_coll = true;
                                    }

                                    if (!alert_coll)
                                    {
                                        foreach (Dictionary<string, string> dipen in dipendenti)
                                        {
                                            if (dipen["CODICE FISCALE"] == codice_fiscale.ToUpper())
                                            {
                                                riga_out.Add("1");            // COLONNA O - ALERT COLLABORATORE     
                                                riga_out.Add(dipen["INIZIO COLLABORAZIONE"]);       // COLONNA P - DATA INIZIO COLLABO
                                                riga_out.Add(dipen["CODICE FISCALE"]);       // COLONNA Q - CF COLLABORATORE
                                                riga_out.Add("");             // COLONNA R - NOTE

                                                if (dipen.ContainsKey("NUMERO ISCRIZIONE"))
                                                {
                                                    riga_out.Add(dipen["NUMERO ISCRIZIONE"]);   // COLONNA S - N ISCRIZ COLLABORATORE
                                                    riga_out.Add(dipen["INIZIO COLLABORAZIONE"]);   // COLONNA T - DATA ISCR DIR COLLABO
                                                    riga_out.Add("ISCRITTO"); // COLONNA U - STATO OAM COLL
                                                    //riga_out.Add(dipen["STATO"]); // COLONNA U - STATO OAM COLL
                                                }
                                                else
                                                {
                                                    riga_out.Add("");       // COLONNA S - N ISCRIZ COLLABORATORE
                                                    riga_out.Add("");       // COLONNA T - DATA ISCR DIR COLLABO
                                                    riga_out.Add("NON ISCRITTO");   // COLONNA U - STATO OAM COLL
                                                }

                                                alert_coll = true;
                                                break;
                                            }
                                        }
                                    }

                                    if (!alert_coll)
                                    {
                                        riga_out.Add("0");      // COLONNA O - ALERT COLLABORATORE
                                        riga_out.Add("");       // COLONNA P - DATA INIZIO COLLABO
                                        riga_out.Add("");       // COLONNA Q - CF COLLABORATORE
                                        riga_out.Add("");       // COLONNA R - NOTE
                                        riga_out.Add("");       // COLONNA S - N ISCRIZ COLLABORATORE
                                        riga_out.Add("");       // COLONNA T - DATA ISCR DIR COLLABO
                                        riga_out.Add("NON ISCRITTO");   // COLONNA U - STATO OAM COLL
                                    }

                                    if (dettagli_oam.ContainsKey("DENOMINAZIONE"))
                                    {
                                        riga_out.Add(dettagli_oam["DENOMINAZIONE"]);  // COLONNA V - DENOM SOC COLL
                                    }
                                    else
                                    {
                                        riga_out.Add(dettagli_oam["COGNOME E NOME"]);
                                    }

                                    riga_out.Add("1");              // COLONNA W - ALERT SOC COLL
                                    riga_out.Add(dettagli_oam["NUMERO ISCRIZIONE"]);  // COLONNA X - N ISCR SOC COLL

                                    if (dettagli_oam["TIPO ELENCO"].StartsWith('A'))
                                    {
                                        riga_out.Add("AGENTE");      // COLONNA Y - TIPO SOC COLL
                                    }
                                    else
                                    {
                                        riga_out.Add("MEDIATORE");   // COLONNA Y - TIPO SOC COLL
                                        mediatore = true;
                                    }

                                    riga_out.Add(dettagli_oam["STATO"]);   // COLONNA Z - STATO SOC COLL


                                    if (piva_soc.Trim() == piva_soc_old.Trim() && piva_intermed.Trim() == piva_intermed_old.Trim())
                                    {
                                        foreach (string s in intermed_prec)
                                        {
                                            riga_out.Add(s);
                                        }
                                    }
                                    else
                                    {
                                        intermed_prec = new List<string>();

                                        bool alert_intermed = false;
                                        List<string> lista_intermediari = new List<string>();
                                        for (int indice_intermed = 0; indice_intermed < 3; indice_intermed++)
                                        {
                                            List<List<string>> prodotti = new List<List<string>>();
                                            if (indice_intermed < mandati.Count)
                                            {
                                                prodotti = OAM.cattura_prodotti(indice_intermed, "Diretti");

                                                if (mandati[indice_intermed]["CODICE FISCALE"] == piva_intermed)
                                                {
                                                    riga_out.Add(mandati[indice_intermed]["DENOMINAZIONE"]);  // COLONNA AA - DENOM INTERMED
                                                    riga_out.Add("1");                          // COLONNA AB - ALERT INTERMED
                                                    riga_out.Add(mandati[indice_intermed]["INIZIO MANDATO"]);  // COLONNA AC - DATA INIZIO INTERMED
                                                    riga_out.Add("");                           // COLONNA AD - N ISCR INTERMED
                                                    riga_out.Add("");                           // COLONNA AE - SEZIONE INTERMED.
                                                    alert_intermed = true;

                                                    intermed_prec.Add(mandati[indice_intermed]["DENOMINAZIONE"]);  // COLONNA AA - DENOM INTERMED
                                                    intermed_prec.Add("1");                          // COLONNA AB - ALERT INTERMED
                                                    intermed_prec.Add(mandati[indice_intermed]["INIZIO MANDATO"]);  // COLONNA AC - DATA INIZIO INTERMED
                                                    intermed_prec.Add("");                           // COLONNA AD - N ISCR INTERMED
                                                    intermed_prec.Add("");                           // COLONNA AE - SEZIONE INTERMED.

                                                    List<string> lista_prodotti = new List<string>();
                                                    for (int indice_prod = 0; indice_prod < 10; indice_prod++)
                                                    {
                                                        if (indice_prod < prodotti.Count)
                                                        {
                                                            if (prodotti[indice_prod][0] == "A.10")
                                                            {
                                                                riga_out.Add(prodotti[indice_prod][0] + " " + prodotti[indice_prod][1]);
                                                                intermed_prec.Add(prodotti[indice_prod][0] + " " + prodotti[indice_prod][1]);
                                                            }
                                                            else
                                                            {
                                                                lista_prodotti.Add(prodotti[indice_prod][0] + " " + prodotti[indice_prod][1]);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            lista_prodotti.Add("");
                                                        }
                                                    }
                                                    foreach (string prod in lista_prodotti)
                                                    {
                                                        riga_out.Add(prod);
                                                        intermed_prec.Add(prod);
                                                    }
                                                }
                                                else
                                                {
                                                    lista_intermediari.Add(mandati[indice_intermed]["DENOMINAZIONE"]);  // COLONNA AA - DENOM INTERMED
                                                    lista_intermediari.Add(mandati[indice_intermed]["INIZIO MANDATO"]);  // COLONNA AC - DATA INIZIO INTERMED
                                                    lista_intermediari.Add("");                           // COLONNA AD - N ISCR INTERMED
                                                    lista_intermediari.Add("");                           // COLONNA AE - SEZIONE INTERMED

                                                    List<string> lista_prodotti = new List<string>();
                                                    for (int indice_prod = 0; indice_prod < 10; indice_prod++)
                                                    {
                                                        if (indice_prod < prodotti.Count)
                                                        {
                                                            if (prodotti[indice_prod][0] == "A.10")
                                                            {
                                                                lista_intermediari.Add(prodotti[indice_prod][0] + " " + prodotti[indice_prod][1]);
                                                            }
                                                            else
                                                            {
                                                                lista_prodotti.Add(prodotti[indice_prod][0] + " " + prodotti[indice_prod][1]);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            lista_prodotti.Add("");
                                                        }
                                                    }

                                                    foreach (string prod in lista_prodotti)
                                                    {
                                                        lista_intermediari.Add(prod);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                for (int i = 0; i < 14; i++)
                                                {
                                                    lista_intermediari.Add("");
                                                }
                                            }
                                        }

                                        for (int i = 0; i < lista_intermediari.Count; i++)
                                        {
                                            if (i == 1)
                                            {
                                                if (!mediatore)
                                                {
                                                    if (!alert_intermed)
                                                    {
                                                        riga_out.Add("0");
                                                        intermed_prec.Add("0");
                                                    }
                                                }
                                                else
                                                {
                                                    if (alert_coll)
                                                    {
                                                        riga_out.Add("1");
                                                        intermed_prec.Add("1");
                                                    }
                                                }
                                            }
                                            riga_out.Add(lista_intermediari[i]);
                                            intermed_prec.Add(lista_intermediari[i]);
                                        }

                                        mediatore = false;
                                    }
                                }
                                else
                                {
                                    riga_out.Add("0");              // COLONNA O - ALERT COLLABORATORE
                                    riga_out.Add("");
                                    riga_out.Add("");
                                    riga_out.Add("");
                                    riga_out.Add("");
                                    riga_out.Add("");
                                    riga_out.Add("NON ISCRITTO");

                                    for (int i = 0; i < 48; i++)
                                    {
                                        riga_out.Add("");
                                    }
                                }
                                //goto scrivi;
                                break;
                        }
                    ivass:
                        browser.Navigate().GoToUrl("https://www.google.it/");
                        browser.Navigate().GoToUrl("https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                        Ivass.collega_driver(browser);

                        string nome_rui = nome;
                        while (true)
                        {
                            Dictionary<string, string> sogg_dett = Ivass.ricerca("", data_nascita, nome_rui, cognome);
                            if (sogg_dett.Count > 0)
                            {
                                riga_out.Add("1");
                                riga_out.Add(sogg_dett["SEZIONE"]);

                                if (sogg_dett.ContainsKey("QUALIFICA DI ESERCIZIO"))
                                {
                                    if (riga_out[1] == "0")
                                    {
                                        List<string> lista_intermed = sogg_dett["QUALIFICA DI ESERCIZIO"].Split('\n').ToList();
                                        foreach (string intermed in lista_intermed)
                                        {
                                            string tmp = string.Empty;
                                            if (intermed.Contains("Responsabile"))
                                            {
                                                tmp = Utils.ExtractBetween(intermed, "società", "iscritto");
                                            }
                                            else if (intermed.Contains("Dipendente") | intermed.Contains("Collaboratore"))
                                            {
                                                tmp = Utils.ExtractBetween(intermed, "intermediario", "iscritto");
                                            }
                                            else if (intermed.Contains("Addetto"))
                                            {
                                                tmp = Utils.ExtractBetween(intermed, "iscritta", "iscritto");
                                            }

                                            if (Utils.RimuoviNDG(tmp) == Utils.RimuoviNDG(intermediario))
                                            {
                                                riga_out[1] = "1";
                                                riga_out[4] = "Dipendente da IVASS";
                                            }
                                        }
                                    }
                                }

                                for (int indice_intermed = 0; indice_intermed < 9; indice_intermed++)
                                {
                                    List<string> lista_intermed = new List<string>();
                                    try
                                    {
                                        lista_intermed = sogg_dett["INTERMEDIARI PER CUI OPERA"].Split('\n').ToList();
                                    }
                                    catch { }

                                    if (indice_intermed < lista_intermed.Count)
                                    {
                                        riga_out.Add(Utils.ExtractBetween(lista_intermed[indice_intermed], "", "Sezione"));
                                        riga_out.Add(Utils.ExtractBetween(lista_intermed[indice_intermed], "Sezione", "Numero"));
                                    }
                                    else
                                    {
                                        riga_out.Add("");
                                        riga_out.Add("");
                                    }
                                }
                            }
                            else
                            {
                                if (nome_rui.Contains(" "))
                                {
                                    nome_rui = Utils.sina(nome_rui, " ");
                                    continue;
                                }
                                else
                                {
                                    riga_out.Add("0");
                                }
                            }
                            break;
                        }
                    //browser.Close();

                    scrivi:
                        //excel.WriteExcel(riga_out, riga, 90, foglio_excel);
                        excel.WriteExcel(riga_out, riga, 14, foglio_excel);
                        riga_out = new List<string>();

                        riga += 1;
                        piva_soc_old = piva_soc;
                        piva_intermed_old = piva_intermed;
                        //riga = 808;
                    }

                    foglio_excel -= 1;
                    //riga = 2;
                }
                excel.Close();
                MsgBox.Show("Elaborazione Terminata");
            }
            catch (Exception ex)
            {
                HTML.chiudiDriver();
                excel.Close();
                log.Write(ex.Message, 4, ex);
                win.Chiudi();
            }
        }
        public static string pulisci_ragsoc(string denominazione)
        {
            denominazione = denominazione.Replace("!", "");
            denominazione = denominazione.Replace("|", "");

            denominazione = denominazione.Replace(" ", "");

            return denominazione;
        }
    }
}
