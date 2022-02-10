using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Collections;

namespace PlanilhaUnico
{
    class Program
    {

        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Planilha-unico";
        static void Main(string[] args)
        {

            void Salvar(IList<object> dados)
            {
                using (SqlConnection connection = new SqlConnection("Nome do Servidor | datasource"))
                {
                    {
                        string sql = "INSERT INTO Lojas (MacroCanal) VALUES (@dados0)"; 
                        using (SqlCommand cmd = new SqlCommand(sql, connection))
                        {
                            connection.Open();
                            cmd.Parameters.Add("@dados0", SqlDbType.VarChar, 255).Value = dados[0].ToString();
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }



                }
            }

            UserCredential credential;


            using (var stream =
                new FileStream("C:/Planilha-Unico1/PlanilhaUnico-1/PlanilhaUnico-1/client_secret_254177701680-1equ9tertind61egl7lv1dkoda0bj466.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
            {

                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = "1QR7jwzYaa9EMzfrdSEUvA8geFrNWFvbBPIN7XPt00cQ";
            String range = "Página1!A1:BE28";                                         // explicitando de onde queremos pegar as informações da planilha para salvar no banco.
            SpreadsheetsResource.ValuesResource.GetRequest request =
            service.Spreadsheets.Values.Get(spreadsheetId, range);


        https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;


            if (values != null && values.Count > 0)
            {

                foreach (var row in values)
                {
                    Console.WriteLine("{0}, | {1}, | {2}, | {3}, | {4}", row[0], row[1], row[3], row[4], row[5]);
                    Salvar(row);

                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();
        }
    }



}

