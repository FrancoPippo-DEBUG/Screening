using DLL;
using OpenQA.Selenium.Chrome;
using System.Globalization;

namespace AIG
{
    internal class AIG
    {
        [STAThread]
        static void Main()
        {
            Win win = new Win("AIG");
            ChromeDriver? browser = null;

            LogFile log = new LogFile();
            Excel excel = new Excel();

            InputBox inp = new InputBox("Riga inizio");
            if (!inp.ret)
            {
                return;
            }

            int riga = 0;
            try
            {
                riga = int.Parse(inp.result);
            }
            catch (Exception)
            {
                riga = 2;
            }

            MessageFrame msg = new MessageFrame(1);

            try
            {
                //browser = html.ApriBrowser_CR("https://servizi.ivass.it/RuirPubblica/");
                browser = HTML.ApriBrowser_CR("https://ruipubblico.ivass.it/rui-pubblica/ng/#/workspace/registro-unico-intermediari");
                Ivass.collega_driver(browser);

                List<string> riga_out = new List<string>();
                string[] riga_letta = new string[] { };
                excel.Open(Utils.CurDir() + @"\AIG.xlsx");
                while (true)
                {
                    log.Write("===========================================================");
                    log.Write("Inizio lavorazione riga " + riga);

                    msg.scrivi("Elaboro riga " + riga);
                    riga_letta = excel.ReadExcel(riga, 1);

                    if (riga_letta[0] == "")
                    {
                        break;
                    }
                    string denominazione = Utils.RimuoviNDG(riga_letta[1]);
                    string indirizzo = riga_letta[2];
                    string cap = riga_letta[3];
                    string citta = riga_letta[4];
                    string provincia = riga_letta[5];

                    string codice_fiscale = riga_letta[10];

                    string indirizzo_completo = string.Format("{0} - {1} {2} ({3})", indirizzo, cap, citta, provincia);
                    string note = string.Empty;
                    string cognome = string.Empty;
                    string nome = string.Empty;
                    string data_nascita = string.Empty;
                    string sesso = string.Empty;

                    string numero_iscrizione = riga_letta[14];
                    Dictionary<string, string> dettaglio_ivass = new Dictionary<string, string>();

                    if (numero_iscrizione != "")
                    {
                        dettaglio_ivass = Ivass.ricerca(numero_iscrizione.Trim());
                    }
                    else
                    {
                        log.Write("Ricerco per denominazione " + denominazione);
                        if (codice_fiscale.Length > 11)
                        {
                            //denominazione = "D'ANGELO NINO";
                            //codice_fiscale = "DNGNNI";
                            Utils.risolviCF(codice_fiscale, denominazione, out cognome, out nome, out data_nascita, out sesso);
                            if (nome != string.Empty)
                            {
                                data_nascita = DateTime.Parse(data_nascita, new CultureInfo("it-IT")).ToString("dd MMMM yyyy");
                                dettaglio_ivass = Ivass.ricerca("", data_nascita, nome, cognome);
                            }
                            else
                            {
                                note = "Aggiungere Numero iscrizione";
                            }
                        }
                        else
                        {
                            dettaglio_ivass = Ivass.ricerca("", citta, denominazione);
                        }
                    }

                    if (dettaglio_ivass.Count > 0)
                    {
                        riga_out.Add(dettaglio_ivass["SEZIONE"].Substring(0, 1)); // SEZIONE
                        riga_out.Add("'" + dettaglio_ivass["DATA ISCRIZIONE"]);

                        if (codice_fiscale.Length < 16)
                        { // PERSONA GIURIDICA
                            List<string> lista = new List<string>();
                            if (dettaglio_ivass.ContainsKey("IMPRESE PER LE QUALI È SVOLTA L'ATTIVITÀ"))
                            {
                                lista = dettaglio_ivass["IMPRESE PER LE QUALI È SVOLTA L'ATTIVITÀ"].Split('\n').ToList();
                            }
                            for (int i = 0; i < 9; i++)
                            {
                                if (lista.Count > i)
                                {
                                    riga_out.Add(lista[i].Trim());
                                }
                                else
                                {
                                    riga_out.Add("");
                                }
                            }
                            string tmp = string.Empty;
                            if (dettaglio_ivass.ContainsKey("QUALIFICA DI ESERCIZIO"))
                            {
                                tmp = dettaglio_ivass["QUALIFICA DI ESERCIZIO"];
                            }
                            riga_out.Add(tmp);

                            if (Utils.RimuoviNDG(dettaglio_ivass["RAGIONE O DENOMINAZIONE SOCIALE"]) != Utils.RimuoviNDG(denominazione))
                            {
                                log.Write("ATTENZIONE: denominazione diversa");
                                note += "[!] DENOMINAZIONE DIVERSA: " + dettaglio_ivass["RAGIONE O DENOMINAZIONE SOCIALE"] + "\n";
                            }
                            riga_out.Add(note.Trim());
                        }
                        else
                        { // PERSONA FISICA
                            for (int i = 0; i < 9; i++)
                            {
                                riga_out.Add("");
                            }
                            // QUALIFICA DI ESERCIZIO
                            string tmp = string.Empty;
                            if (dettaglio_ivass.ContainsKey("QUALIFICA DI ESERCIZIO"))
                            {
                                tmp = dettaglio_ivass["QUALIFICA DI ESERCIZIO"];
                            }
                            riga_out.Add(tmp);
                        }
                    }
                    else
                    {
                        log.Write("La ricerca non ha prodotto risultati");
                        for (int i = 0; i < 12; i++)
                        {
                            riga_out.Add("");
                        }
                        if (note == string.Empty)
                        {
                            riga_out.Add("[X] NON ISCRITTO AL RUI");
                        }
                        else
                        {
                            riga_out.Add(note);
                        }
                    }

                    excel.WriteExcel(riga_out, riga, 16);
                    riga_out = new List<string>();
                    riga += 1;
                }
                excel.Close();
                browser.Close();
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
    }
}
