using DLL;
using OpenQA.Selenium.Chrome;
using System.Net.Sockets;

namespace CompassRUI
{
    internal class CompassRUI
    {
        [STAThread]
        static void Main()
        {
            Win win = new Win("Compass RUI");
            ChromeDriver browser = null;

            LogFile log = new LogFile();
            LogFile repo = new LogFile("REPORT_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log");
            Excel excel = new Excel();

            try
            {
                string[] riga_letta = new string[] { };
                Dictionary<string, string> result = new Dictionary<string, string>();
                List<string> riga_out = new List<string>();

                string[] natura_giuridica = new string[] { "SRLS", "S.R.L.S.", "SRL", "S.R.L.", " SAS", "S.A.S.", "SNC", "S.N.C", " SPA", "S.P.A.",
                                                           "SOCIETA IN NOME COLLETTIVO", "SOCIETA IN ACCOMANDITA SEMPLICE", "SOCIETA PER AZIONI",
                                                           "SOCIETA' IN NOME COLLETTIVO", "SOCIETA' IN ACCOMANDITA SEMPLICE", "SOCIETA' PER AZIONI",
                                                           "SOCIETÀ IN NOME COLLETTIVO", "SOCIETÀ IN ACCOMANDITA SEMPLICE", "SOCIETÀ PER AZIONI",
                                                           "SOCIETA' A RESPONSABILITA' LIMITATA", "SOCIETA A RESPONSABILITA LIMITATA", "SOCIETÀ A RESPONSABILITÀ LIMITATA",
                                                           "SOCIETA' A RESPONSABILITA' LIMITATA SEMPLIFICA", "SOCIETA A RESPONSABILITA LIMITATA SEMPLIFICA", "SOCIETÀ A RESPONSABILITÀ LIMITATA SEMPLIFICA"};

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

                browser = HTML.ApriBrowser("CR", "https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                Ivass.collega_driver(browser);

                excel.Open(Utils.CurDir() + "\\IVASS.xlsx");
                MessageFrame msg = new MessageFrame(1);
                while (true)
                {
                    bool din = false;
                    log.Write("----------- Elaboro riga " + riga);
                    msg.scrivi("Elaboro riga " + riga);
                    riga_letta = excel.ReadExcel(riga);
                    
                    if (riga_letta[65] != "")
                    {
                        /*
                        if (!riga_letta[7].Contains("COMPASS"))
                        {
                            riga_out = ordinaIntermed(riga_letta.ToList());
                            excel.WriteExcel(riga_out, riga);
                            riga_out = [];
                        }
                        */
                        riga += 1;
                        continue;
                    }
                    //riga_letta = ut.ReadExcel(ut.CurDir() + "\\IVASS.xlsx", riga);
                    if (riga_letta[0] == "")
                    {
                        break;
                    }
                    string tipo_intermediario = riga_letta[0];

                    string denominazione = riga_letta[2];
                    if (denominazione.Contains(" DIN"))
                    {
                        din = true;
                    }
                    string indirizzo = riga_letta[3];
                    string comune = riga_letta[4];
                    if (comune.EndsWith('\''))
                    {
                        comune = comune.Remove(comune.Length - 1, 1);
                    }
                    string provincia = riga_letta[5];
                    string agente = riga_letta[7];
                    string numero_iscr = riga_letta[64].Trim();

                    if (numero_iscr.Contains("NUMERO ISCRIZIONE"))
                    {
                        numero_iscr = numero_iscr.Replace("NUMERO ISCRIZIONE", "").Trim();
                        numero_iscr = numero_iscr.Replace(":", "").Trim();
                    }
                    else if (numero_iscr.Contains("NUM ISC"))
                    {
                        numero_iscr = numero_iscr.Replace("NUM ISC", "").Trim();
                    }
                    else if (numero_iscr.Contains("IVASS"))
                    {
                        numero_iscr = numero_iscr.Replace("IVASS", "").Trim();
                    }
                    else if (numero_iscr.Contains("E PLAFOND"))
                    {
                        numero_iscr = numero_iscr.Replace("E PLAFOND", "").Trim();
                    }

                    int indice_sogg = 1;
                    List<string> soggetto = new List<string>();
                    List<List<string>> soggetti = new List<List<string>>();
                    while (indice_sogg <= 56)
                    {
                        if (riga_letta[indice_sogg + 7] == "")
                        {
                            break;
                        }
                        else
                        {
                            if (riga_letta[indice_sogg + 7].Contains("'") &&
                                riga_letta[indice_sogg + 7].IndexOf("'") < 4)
                            {
                                riga_letta[indice_sogg + 7] = riga_letta[indice_sogg + 7].Replace("' ", "'");
                            }

                            soggetto.Add(riga_letta[indice_sogg + 7]);

                            if (indice_sogg % 4 == 0)
                            {
                                int trovati = 0;
                                foreach (string parola in Utils.dividi_parole(soggetto[0]))
                                {
                                    if (denominazione.Contains(parola))
                                    {
                                        trovati += 1;
                                    }
                                }

                                if (trovati >= 2)
                                {
                                    din = true;
                                    log.Write("Tratto come DIN perchè " + soggetto[0] + " in " + denominazione);
                                }
                                soggetto[0] = Utils.TogliApostrofi(soggetto[0]);
                                soggetti.Add(soggetto);
                                soggetto = new List<string>();
                            }
                            indice_sogg += 1;
                        }
                    }
                    // ================================================== TIPO INTERMEDIARIO: CV
                    if (tipo_intermediario.Trim() == "CV" || tipo_intermediario.Trim() == "SA")
                    {
                        bool societa = false;
                        foreach (string ndg in natura_giuridica)
                        {
                            if (denominazione.Contains(ndg))
                            {
                                denominazione = denominazione.Remove(denominazione.IndexOf(ndg), ndg.Length);
                                societa = true;
                                log.Write("Tratto come SOCIETA perchè " + ndg + " in " + denominazione);
                            }
                        }
                        if (denominazione.EndsWith("SAS"))
                        {
                            societa = true;
                            log.Write("Tratto come SOCIETA perchè SAS in " + denominazione);
                        }
                        if (denominazione.Contains("'"))
                        {
                            if (denominazione.IndexOf("'") < 4)
                            {
                                denominazione = denominazione.Replace("' ", "'");
                            }
                        }

                        if (denominazione.Contains(" DIN") ||
                            (din && !societa))
                        {
                            denominazione = denominazione.Replace(" DIN", "");
                            if (Utils.quante_parole(denominazione) >= 2)
                            {
                                if (soggetti.Count > 0)
                                {
                                    for (int i = 0; i < 14; i++)
                                    {
                                        if (i <= soggetti.Count - 1)
                                        {
                                            string cognome = Utils.ExtractBetween(soggetti[i][0], "", " ").Trim();
                                            if (cognome.Length < 4 && Utils.quante_parole(soggetti[i][0]) > 2)
                                            {
                                                cognome += " " + Utils.ExtractBetween(Utils.desa(soggetti[i][0], cognome).Trim(), "", " ").Trim();
                                                repo.Write("RIGA: " + riga + " ################### DIN SOGGETTO " + (i + 1) + " TROPPI NOMI - " + soggetti[i][0]);
                                            }
                                            if (cognome.EndsWith('\''))
                                            {
                                                cognome = cognome.Remove(cognome.Length - 1, 1);
                                            }

                                            string nome = soggetti[i][0].Replace(cognome, "").Trim();
                                            if (nome.Contains(" "))
                                            {
                                                //nome = nome.Substring(0, nome.IndexOf(" ")); PRIMO NOME
                                                nome = Utils.desa(nome, " "); // ULTIMO NOME
                                                // AVVISO DENOMINAZIONE CON PIU DI DUE PAROLE
                                                repo.Write("RIGA: " + riga + " ################### DIN SOGGETTO " + (i + 1) + " TROPPI NOMI - " + soggetti[i][0]);
                                            }
                                            log.Write("Ho separato nome: " + nome + " e cognome: " + cognome);

                                            if (!(denominazione.Contains(cognome) || denominazione.Contains(nome)))
                                            {
                                                continue;
                                            }

                                            string nato_il = DateTime.Parse(soggetti[i][1]).ToString("dd/MM/yyyy");
                                            string nato_a = soggetti[i][2];

                                            //List<List<string>> sogg_dett = new List<List<string>>();
                                            Dictionary<string, string> sogg_dett = new Dictionary<string, string>();

                                            if (numero_iscr != "")
                                            {
                                                sogg_dett = Ivass.ricerca(numero_iscr, nato_il, nome, cognome);
                                            }
                                            else
                                            {
                                                sogg_dett = Ivass.ricerca("", nato_il, nome, cognome);

                                                if (sogg_dett.Count > 0)
                                                {
                                                    result = sogg_dett;
                                                }
                                            }
                                            if (din && result.Count == 0)
                                            {
                                                result = sogg_dett;
                                            }
                                            //result = sogg_dett;
                                        }

                                        if (result.Count > 0)
                                        {
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    riga_out.Add("INSERIRE NUMERO ISCRIZIONE");
                                    excel.WriteExcel(riga_out, riga, 66);
                                    riga_out = new List<string>();
                                    riga += 1;
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            result = Ivass.ricerca(numero_iscr, comune, denominazione);

                            if (result.Count == 0)
                            { // PROVO A CERCARE LA SOCIETA' PARTENDO DAI SOGGETTI
                                foreach (List<string> x in soggetti)
                                {
                                    string cognome = Utils.ExtractBetween(x[0], "", " ").Trim();
                                    if (cognome.Length < 4 && Utils.quante_parole(x[0]) > 2)
                                    {
                                        cognome += " " + Utils.ExtractBetween(Utils.desa(x[0], cognome).Trim(), "", " ").Trim();
                                        repo.Write("RIGA: " + riga + " ################### SOGG SOC " + (soggetti.IndexOf(x) + 1) + " TROPPI NOMI - " + x[0]);
                                    }
                                    if (cognome.EndsWith('\''))
                                    {
                                        cognome = cognome.Remove(cognome.Length - 1, 1);
                                    }
                                    string nome = x[0].Replace(cognome, "").Trim();
                                    if (nome.Contains(" "))
                                    {
                                        //nome = nome.Substring(0, nome.IndexOf(" ")); PRIMO NOME
                                        nome = Utils.desa(nome, " "); // ULTIMO NOME
                                        // AVVISO DENOMINAZIONE CON PIU DI DUE PAROLE
                                        repo.Write("RIGA: " + riga + " ################### SOGG SOC " + (soggetti.IndexOf(x) + 1) + " TROPPI NOMI - " + x[0]);
                                    }
                                    string nato_il = DateTime.Parse(x[1]).ToString("dd/MM/yyyy");
                                    string nato_a = x[2];

                                    Dictionary<string, string> sogg_dett = Ivass.ricerca("", nato_il, nome, cognome);
                                    if (sogg_dett.Count > 0)
                                    {
                                        if (sogg_dett.Count > 0)
                                        {
                                            if (sogg_dett.ContainsKey("QUALIFICA DI ESERCIZIO"))
                                            {
                                                List<string> lista_qualifica = sogg_dett["QUALIFICA DI ESERCIZIO"].Split('\n').ToList();

                                                foreach (string tmp in lista_qualifica)
                                                {
                                                    if (tmp.Contains("Responsabile di società"))
                                                    {
                                                        denominazione = Utils.ExtractBetween(tmp, "Responsabile di società", "iscritto/a");
                                                        result = Ivass.ricerca("", comune, denominazione);                                                        
                                                    }
                                                    else if (tmp.Contains("Responsabile dell'attività di intermediazione della soc"))
                                                    {
                                                        denominazione = Utils.ExtractBetween(tmp, "Responsabile dell'attività di intermediazione della soc.", " / ");
                                                        result = Ivass.ricerca("", comune, denominazione);
                                                    }
                                                    else if (tmp.Contains("Collaboratore di intermediario"))
                                                    {
                                                        denominazione = Utils.ExtractBetween(tmp, "Collaboratore di intermediario", "iscritto/a");
                                                        result = Ivass.ricerca("", comune, denominazione);
                                                    }
                                                    else if (tmp.Contains("Addetto di società iscritta"))
                                                    {
                                                        denominazione = Utils.ExtractBetween(tmp, "Addetto di società iscritta", "iscritto/a");
                                                        result = Ivass.ricerca("", comune, denominazione);
                                                    }
                                                    else if (tmp.Contains("Dipendente dell'intermediario"))
                                                    {
                                                        denominazione = Utils.ExtractBetween(tmp, "Dipendente dell'intermediario", "iscritto/a");
                                                        result = Ivass.ricerca("", comune, denominazione);
                                                    }

                                                    if (result.Count == 0)
                                                    {
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        break;
                                                    }
                                                }
                                            }
                                            else if (sogg_dett.ContainsKey("CARICHE SOCIETARIE"))
                                            {
                                                List<string> lista_qualifica = sogg_dett["CARICHE SOCIETARIE"].Split('\n').ToList();

                                                foreach (string tmp in lista_qualifica)
                                                {
                                                    if (tmp.Contains("Responsabile di società"))
                                                    {
                                                        denominazione = Utils.ExtractBetween(tmp, "Responsabile di società", "iscritto/a");
                                                        result = Ivass.ricerca("", comune, denominazione);
                                                    }
                                                    else if (tmp.Contains("Responsabile dell'attività di intermediazione per"))
                                                    {
                                                        denominazione = Utils.desa(tmp, "intermediazione per");
                                                        denominazione = Utils.ExtractBetween(denominazione, " ", "dal");
                                                        result = Ivass.ricerca("", comune, denominazione);
                                                    }

                                                    if (result.Count > 0)
                                                    {
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (result.Count > 0)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        if (result.Count > 0)
                        {
                            string note = string.Empty;

                            string tmp = string.Empty;
                            if (result.ContainsKey("RAGIONE O DENOMINAZIONE SOCIALE"))
                            {
                                tmp = result["RAGIONE O DENOMINAZIONE SOCIALE"];
                            }
                            else
                            {
                                tmp = result["NOMINATIVO"];
                            }

                            if (pulisci_ragsoc(denominazione.Trim(), natura_giuridica) !=
                                pulisci_ragsoc(tmp, natura_giuridica))
                            {
                                note += "DENOM: " + tmp + "\r\n";
                            }

                            if (result.ContainsKey("OPERATIVITÀ"))
                            {
                                if (result["OPERATIVITÀ"].Contains("INOPERATIVO"))
                                {
                                    note += result["OPERATIVITÀ"];
                                }
                                else
                                {
                                    if (result.ContainsKey("OPERATIVITÀ INDIVIDUALE"))
                                    {
                                        if (result["OPERATIVITÀ INDIVIDUALE"].Contains("INOPERATIVO"))
                                        {
                                            note += result["OPERATIVITÀ INDIVIDUALE"] + "\r\n";
                                        }
                                    }

                                    if (result.ContainsKey("OPERATIVITÀ SOCIETARIA"))
                                    {
                                        if (result["OPERATIVITÀ SOCIETARIA"].Contains("INOPERATIVO"))
                                        {
                                            note += result["OPERATIVITÀ SOCIETARIA"] + "\r\n";
                                        }
                                    }
                                }
                            }

                            riga_out.Add("'" + DateTime.Now.ToString("dd/MM/yyyy"));
                            riga_out.Add("1");
                            riga_out.Add(result["SEZIONE"]);
                            riga_out.Add(result["NUMERO ISCRIZIONE"]);
                            riga_out.Add("ISCRITTO");
                            try
                            {
                                riga_out.Add("'" + result["DATA ISCRIZIONE"]);
                            }
                            catch { riga_out.Add(""); }

                            riga_out.Add(note.Trim());
                            // ------------------------------------------------------------- SOC_COLL
                            for (int i = 0; i < 14; i++)
                            {
                                if (i <= soggetti.Count - 1)
                                {
                                    string cognome = Utils.ExtractBetween(soggetti[i][0], "", " ").Trim();
                                    if (cognome.Length < 4 && Utils.quante_parole(soggetti[i][0]) > 2)
                                    {
                                        cognome += " " + Utils.ExtractBetween(Utils.desa(soggetti[i][0], cognome).Trim(), "", " ").Trim();
                                    }
                                    string nome = soggetti[i][0].Replace(cognome, "").Trim();
                                    if (nome.Contains(" "))
                                    {
                                        //nome = nome.Substring(0, nome.IndexOf(" ")); PRIMO NOME
                                        nome = Utils.desa(nome, " "); // ULTIMO NOME
                                    }
                                    string nato_il = DateTime.Parse(soggetti[i][1]).ToString("dd/MM/yyyy");
                                    string nato_a = soggetti[i][2];

                                    Dictionary<string, string> sogg_dett = Ivass.ricerca("", nato_il, nome, cognome);
                                    if (sogg_dett.Count > 0)
                                    {
                                        riga_out.Add("1");
                                    }
                                    else
                                    {
                                        riga_out.Add("0");
                                    }
                                }
                                else
                                {
                                    riga_out.Add("");
                                }
                            }
                            // ----------------------------------------------------- INTERMEDIARI
                            List<string> lista_intermediari = new List<string>();
                            List<string> intermed = new List<string>();
                            if (result.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                            {
                                intermed = result["INTERMEDIARI PER CUI OPERA"].Split('\n').ToList();
                            }
                            else if (result.ContainsKey("IMPRESE PER LE QUALI È SVOLTA L'ATTIVITÀ"))
                            {
                                intermed = result["IMPRESE PER LE QUALI È SVOLTA L'ATTIVITÀ"].Split('\n').ToList();
                            }

                            for (int i = 0; i < 20; i++)
                            {
                                if (i < intermed.Count && intermed.Count > 0)
                                {
                                    if (intermed[i].Contains("Sezione"))
                                    {
                                        intermed[i] = Utils.ExtractBetween(intermed[i], "", "Sezione").Trim();
                                    }
                                    // AGGIUNGO IN RIGA_OUT SOLO GLI INTERMEDIARI CON ALERT = 1
                                    switch (agente)
                                    {
                                        case string x when x.Contains("ALLIANZ"):
                                            if (//intermed[i].Contains("AWP P&C S.A.") |
                                                //intermed[i].Contains("AWP P & C S.A.") |
                                                //intermed[i].Contains("GENIALLOYD") |
                                                intermed[i].Contains("ALLIANZ"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains(" GAMA") ||
                                                           x.Contains("MAGAP"):
                                            if (intermed[i].Contains("ALLIANZ"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("GASAV") ||
                                                           x.Contains("MAGAP") ||
                                                           x.Contains("FATA ASSICURAZIONI") ||
                                                           x == "G.A.A. NAVALE":
                                            if (intermed[i].Contains("UNIPOL") ||
                                                intermed[i].Contains("ALLIANZ") ||
                                                intermed[i].Contains("GENERALI"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;
                                        /*
                                        case string x when x.Contains("FATA ASSICURAZIONI"):
                                            if (intermed[i].Contains("FATA ASSICURAZIONI") | 
                                                intermed[i].Contains("SOCIETA' CATTOLICA") |                                                
                                                intermed[i].Contains("GENERALI"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }                                    
                                            break;
                                        */
                                        case string x when x.Contains("SOCIETA' CATTOLICA"):
                                            if (intermed[i].Contains("SOCIETA' CATTOLICA") ||
                                                intermed[i].Contains("GENERTEL"))
                                            //intermed[i].Contains("FATA ASSICURAZIONI")) 
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("GENERALI ITALIA"):
                                            if (//intermed[i].Contains("D.A.S.") |
                                                //intermed[i].Contains("EUROP ASSISTANCE") |
                                                //intermed[i].Contains("GENERTELLIFE") |
                                                intermed[i].Contains("GENERALI ITALIA"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("G.A.T.E."):
                                            if (intermed[i].Contains("NOBIS"))
                                            //intermed[i].Contains("ERGO") |
                                            //intermed[i].Contains("DARAG") |
                                            //intermed[i].Contains("EUROVITA") |
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("GENERTEL"):
                                            if (intermed[i].Contains("GENERTEL") ||
                                                intermed[i].Contains("GROUPAMA") ||
                                                intermed[i].Contains("GENERALI"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("TUA ASSICURAZIONI"):
                                            //x.Contains("SOCIETA' CATTOLICA"):
                                            if (//intermed[i].Contains("SOCIETA' CATTOLICA") |
                                                intermed[i].Contains("TUA ASSICURAZIONI"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("VITTORIA ASSICURAZIONI"):
                                            if (//intermed[i].Contains("D.A.S.") |
                                                intermed[i].Contains("VITTORIA ASSICURAZIONI"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;
                                        
                                        case string x when x.Contains("GROUPAMA"):
                                            if (intermed[i].Contains("GROUPAMA"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;
                                        
                                        case string x when x.Contains("AXA"):
                                            if (intermed[i].Contains("AXA"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;

                                        case string x when x.Contains("SNAS"):
                                            if (intermed[i].Contains("SNAS"))
                                            {
                                                riga_out.Add("1");
                                                riga_out.Add(intermed[i]);
                                            }
                                            else
                                            {
                                                lista_intermediari.Add("0");
                                                lista_intermediari.Add(intermed[i]);
                                            }
                                            break;
                                        // SE ALERT INTERMEDIARIO = 0 LO AGGINGO ALLA LISTA INTERMEDIARI
                                        default:
                                            lista_intermediari.Add("0");
                                            lista_intermediari.Add(intermed[i]);
                                            break;
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            // INSERISCO INFINE TUTTI GLI INTERMEDIARI CON ALERT = 0
                            foreach (string interm in lista_intermediari)
                            {
                                riga_out.Add(interm);
                            }
                        }
                        else
                        {
                            riga_out.Add("'" + DateTime.Now.ToString("dd/MM/yyyy"));
                            riga_out.Add("0");
                            riga_out.Add("");
                            riga_out.Add("");
                            riga_out.Add("NON ISCRITTO");
                        }
                    }
                    else if (tipo_intermediario == "PV") // ================================================== TIPO INTERMEDIARIO: PV
                    {
                        if (soggetti.Count > 0)
                        {
                            string cognome = Utils.ExtractBetween(soggetti[0][0], "", " ").Trim();
                            if (cognome.Length < 4 && Utils.quante_parole(soggetti[0][0]) > 2)
                            {
                                cognome += " " + Utils.ExtractBetween(Utils.desa(soggetti[0][0], cognome).Trim(), "", " ").Trim();
                            }

                            string nome = soggetti[0][0].Replace(cognome, "").Trim();
                            if (nome.Contains(" "))
                            {
                                //nome = nome.Substring(0, nome.IndexOf(" ")); PRIMO NOME
                                nome = Utils.desa(nome, " "); // ULTIMO NOME
                                // AVVISO DENOMINAZIONE CON PIU DI DUE PAROLE
                            }

                            string nato_il = DateTime.Parse(soggetti[0][1]).ToString("dd/MM/yyyy");
                            string nato_a = soggetti[0][2];

                            Dictionary<string, string> sogg_dett = new Dictionary<string, string>();
                            if (numero_iscr != "")
                            {
                                sogg_dett = Ivass.ricerca(numero_iscr, nato_il, nome, cognome);
                            }
                            else
                            {
                                sogg_dett = Ivass.ricerca("", nato_il, nome, cognome);
                            }

                            if (sogg_dett.Count > 0)
                            {
                                result = sogg_dett;

                                riga_out.Add("'" + DateTime.Now.ToString("dd/MM/yyyy"));
                                riga_out.Add("1");
                                riga_out.Add(result["SEZIONE"]);
                                riga_out.Add(result["NUMERO ISCRIZIONE"]);
                                riga_out.Add("ISCRITTO");
                                riga_out.Add(result["DATA ISCRIZIONE"]);
                                riga_out.Add("");

                                List<string> intermediari = new List<string>();

                                if (result["SEZIONE"].Contains("B -"))
                                {
                                    riga_out.Add("1");
                                }
                                else
                                {
                                    bool trovato = false;
                                    if (result.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                                    {
                                        intermediari = result["INTERMEDIARI PER CUI OPERA"].Split('\n').ToList();
                                    }

                                    foreach (string intermed in intermediari)
                                    {
                                        string tmp = string.Empty;
                                        if (intermed.ToUpper().Contains("SEZIONE"))
                                        {
                                            tmp = intermed.ToUpper().Substring(0, intermed.ToUpper().IndexOf("SEZIONE")).Trim();
                                        }
                                        else
                                        {
                                            tmp = intermed;
                                        }
                                        if (pulisci_ragsoc(tmp, natura_giuridica) ==
                                            pulisci_ragsoc(agente, natura_giuridica))
                                        {
                                            riga_out.Add("1");
                                            trovato = true;
                                            break;
                                        }
                                    }

                                    if (!trovato)
                                    {
                                        riga_out.Add("0");
                                    }
                                }

                                for (int i = 0; i < 13; i++)
                                {
                                    riga_out.Add("");
                                }

                                List<string> lista_intermediari = new List<string>();
                                intermediari = new List<string>();
                                if (result.ContainsKey("INTERMEDIARI PER CUI OPERA"))
                                {
                                    intermediari = result["INTERMEDIARI PER CUI OPERA"].Split('\n').ToList();
                                }

                                for (int i = 0; i < 20; i++)
                                {
                                    if (i < intermediari.Count && intermediari.Count > 0)
                                    {
                                        if (intermediari[i].ToUpper().Contains("SEZIONE"))
                                        {
                                            intermediari[i] = Utils.ExtractBetween(intermediari[i].ToUpper(), "", "SEZIONE").Trim();
                                        }
                                        // AGGIUNGO IN RIGA_OUT SOLO GLI INTERMEDIARI CON ALERT = 1
                                        if (pulisci_ragsoc(intermediari[i], natura_giuridica) ==
                                            pulisci_ragsoc(agente, natura_giuridica))
                                        {
                                            riga_out.Add("1");
                                            riga_out.Add(intermediari[i]);
                                        }
                                        else
                                        {
                                            lista_intermediari.Add("0");
                                            lista_intermediari.Add(intermediari[i]);
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                // INSERISCO INFINE TUTTI GLI INTERMEDIARI CON ALERT = 0
                                if (lista_intermediari.Count > 0)
                                {
                                    foreach (string intermed in lista_intermediari)
                                    {
                                        riga_out.Add(intermed);
                                    }
                                }
                            }
                            else
                            {
                                riga_out.Add("'" + DateTime.Now.ToString("dd/MM/yyyy"));
                                riga_out.Add("0");
                                riga_out.Add("");
                                riga_out.Add("");
                                riga_out.Add("NON ISCRITTO");
                            }
                        }
                    }
                    riga_out = ordinaIntermed(riga_out);
                    excel.WriteExcel(riga_out, riga, 66);
                    //ut.WriteExcel(ut.CurDir() + "\\IVASS.xlsx", riga_out, riga, 66);
                    riga_out = new List<string>();
                    result = new Dictionary<string, string>();
                    riga += 1;
                }


                excel.Close();
                browser.Quit();
                MsgBox.Show("Elaborazione Terminata!");
            }
            catch (Exception ex)
            {
                HTML.chiudiDriver();
                excel.Close();
                log.Write(ex.Message, 4, ex);
                win.Chiudi();
            }
        }

        public static string pulisci_ragsoc(string denominazione, string[] ndg)
        {
            // AGENZIA GENERALE INA ASSITALIAMONZA B.B.R. ASSICURAZIONI
            // AGENZIA GENERALE INA ASSITALIA MONZA B.B.R. ASSICURAZIONI S.R.L.

            foreach (string s in ndg)
            {
                denominazione = denominazione.Replace(s, "");
            }
            denominazione = denominazione.Replace("&", "");
            denominazione = denominazione.Replace(" E ", "");
            denominazione = denominazione.Replace("C.", "");
            denominazione = denominazione.Replace("CO.", "");
            denominazione = denominazione.Replace(".", "");
            denominazione = denominazione.Replace("DIN", "");
            denominazione = denominazione.Replace("'", "");
            denominazione = denominazione.Replace("ASSICURAZIONI", "");
            denominazione = denominazione.Replace("ASSICURAZIONE", "");
            denominazione = denominazione.Replace("ASS. ", "");

            denominazione = denominazione.Replace(" ", "");

            return denominazione;
        }

        public static List<string> ordinaIntermed(List<string> riga_out)
        {
            if (riga_out[0] == "CV")
            {
                string agente = riga_out[7];

                if (agente.Contains("COMPASS"))
                {
                    return riga_out;
                }

                List<string> intermed_1 = [];
                List<string> intermed_0 = [];

                for (int i = 87; i < riga_out.Count; i += 2)
                {
                    string intermed = riga_out[i];
                    if (intermed == "")
                    {
                        break;
                    }

                    switch (agente)
                    {
                        case string x when x.Contains("ALLIANZ"):
                            if (intermed.Contains("AWP P&C S.A.") ||
                                intermed.Contains("AWP P & C S.A.") ||
                                intermed.Contains("GENIALLOYD") ||
                                intermed.Contains("ALLIANZ"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains(" GAMA") ||
                                           x.Contains("MAGAP"):
                            if (intermed.Contains("ALLIANZ"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("GASAV") ||
                                           x.Contains("MAGAP") ||
                                           x.Contains("FATA ASSICURAZIONI") ||
                                           x == "G.A.A. NAVALE":
                            if (intermed.Contains("UNIPOL") ||
                                intermed.Contains("ALLIANZ") ||
                                intermed.Contains("GENERALI"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;
                        
                        case string x when x.Contains("FATA ASSICURAZIONI"):
                            if (intermed.Contains("FATA ASSICURAZIONI") ||
                                intermed.Contains("SOCIETA' CATTOLICA") ||                                                
                                intermed.Contains("GENERALI"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }                                    
                            break;
                        
                        case string x when x.Contains("SOCIETA' CATTOLICA"):
                            if (intermed.Contains("SOCIETA' CATTOLICA") ||
                                intermed.Contains("GENERTEL") ||
                                intermed.Contains("FATA ASSICURAZIONI"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("GENERALI ITALIA"):
                            if (intermed.Contains("D.A.S.") ||
                                intermed.Contains("EUROP ASSISTANCE") ||
                                intermed.Contains("GENERTELLIFE") ||
                                intermed.Contains("GENERALI ITALIA"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("G.A.T.E."):
                            if (intermed.Contains("NOBIS") ||
                                intermed.Contains("ERGO") ||
                                intermed.Contains("DARAG") ||
                                intermed.Contains("EUROVITA"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("GENERTEL"):
                            if (intermed.Contains("GENERTEL") ||
                                intermed.Contains("GROUPAMA") ||
                                intermed.Contains("GENERALI"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("TUA ASSICURAZIONI") ||
                                           x.Contains("SOCIETA' CATTOLICA"):
                            if (intermed.Contains("SOCIETA' CATTOLICA") ||
                                intermed.Contains("TUA ASSICURAZIONI"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("VITTORIA ASSICURAZIONI"):
                            if (intermed.Contains("D.A.S.") ||
                                intermed.Contains("VITTORIA ASSICURAZIONI"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("GROUPAMA"):
                            if (intermed.Contains("GROUPAMA"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("AXA"):
                            if (intermed.Contains("AXA"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;

                        case string x when x.Contains("SNAS"):
                            if (intermed.Contains("SNAS"))
                            {
                                intermed_1.Add(intermed);
                            }
                            else
                            {
                                intermed_0.Add(intermed);
                            }
                            break;
                        // SE ALERT INTERMEDIARIO = 0 LO AGGINGO ALLA LISTA INTERMEDIARI
                        default:
                            intermed_0.Add(intermed);
                            break;
                    }
                }

                int indice_intermed = 0;
                for (int i = 86; i < riga_out.Count; i += 2)
                {
                    if (indice_intermed < intermed_1.Count)
                    {
                        riga_out[i] = "1";
                        riga_out[i + 1] = intermed_1[indice_intermed];
                        indice_intermed += 1;
                        if (indice_intermed == intermed_1.Count)
                        {
                            indice_intermed = 0; // azzero il contatore
                            intermed_1 = []; // svuoto intermed_1
                        }
                        continue;
                    }

                    if (indice_intermed < intermed_0.Count)
                    {
                        riga_out[i] = "0";
                        riga_out[i + 1] = intermed_0[indice_intermed];
                        indice_intermed += 1;
                        if (indice_intermed == intermed_0.Count)
                        {
                            break;
                        }
                        continue;
                    }
                }
            }
            return riga_out;
        }
    }
}
