using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

using System.Data.SqlClient;

namespace EvReceiverCalendar.EventReceiver1
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class EventReceiver1 : SPItemEventReceiver
    {

       /// <summary>
       /// An item is being added.
       /// </summary>
       public override void ItemAdding(SPItemEventProperties properties)
       {
           var assemblyName = System.Reflection.Assembly.GetAssembly(typeof(EvReceiverCalendar.EventReceiver1.EventReceiver1)).FullName;

           var title = properties.List.Title;
           if (title == "Calendario Auto")
           {
                var Title = properties.AfterProperties["Title"].ToString();
                var StartDate = properties.AfterProperties["StartDate"].ToString();
                var _EndDate = properties.AfterProperties["_EndDate"].ToString();
                var StatoAuto = properties.AfterProperties["StatoAuto"].ToString();
                var Destinazione = properties.AfterProperties["Destinazione"].ToString();
                var NrFoglio = properties.AfterProperties["NrFoglio"].ToString();
                var Matricola = properties.AfterProperties["Matricola"].ToString();
                var Data_x0020_Partenza = properties.AfterProperties["Data_x0020_Partenza"].ToString();
                var Data_x0020_Rientro = properties.AfterProperties["Data_x0020_Rientro"].ToString();
                var Commessa = properties.AfterProperties["Commessa"].ToString();
                var Lista = title;

                //string connectionString = @"Password=Sas1A;Persist Security Info=True;User ID=sasa;Initial Catalog=Intranet;Data Source=.\SQLExpress";
                string connectionString = @"Password=vitrociset;Persist Security Info=True;User ID=sa;Initial Catalog=Intranet;Data Source=10.20.100.72";
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    String insertStatement = "INSERT INTO [Intranet].[AutoAziendali].[Prenotazioni] " +
                                                " ([Title] ,[StartDate] ,[_EndDate] ,[StatoAuto] ,[Destinazione] ,[NrFoglio] ,[Matricola] ,[Data_x0020_Partenza] ,[Data_x0020_Rientro] ,[Commessa], [Evento],Origin,Lista) " +
                                                " VALUES (@Title, @StartDate, @_EndDate, @StatoAuto, @Destinazione, @NrFoglio, @Matricola, @Data_x0020_Partenza, @Data_x0020_Rientro, @Commessa, 'ItemAdding', '" + assemblyName + "', '"+title+"')";

                    using (SqlCommand command = new SqlCommand(insertStatement, conn))
                    {
                        command.Parameters.AddWithValue("@Title", Title);
                        command.Parameters.AddWithValue("@StartDate", StartDate);
                        command.Parameters.AddWithValue("@_EndDate", _EndDate);
                        command.Parameters.AddWithValue("@StatoAuto", StatoAuto);
                        command.Parameters.AddWithValue("@Destinazione", Destinazione);
                        command.Parameters.AddWithValue("@NrFoglio", NrFoglio);
                        command.Parameters.AddWithValue("@Matricola", Matricola);
                        command.Parameters.AddWithValue("@Data_x0020_Partenza", Data_x0020_Partenza);
                        command.Parameters.AddWithValue("@Data_x0020_Rientro", Data_x0020_Rientro);
                        command.Parameters.AddWithValue("@Commessa", Commessa);

                        conn.Open();
                        int result = command.ExecuteNonQuery();

                        // Check Error
                        //if (result < 0)
                        //     Console.WriteLine("Error inserting data into Database!");
                    }
                }
           }

           base.ItemAdding(properties);
       }

       /// <summary>
       /// An item is being updated.
       /// </summary>
       public override void ItemUpdating(SPItemEventProperties properties)
       {
           var assemblyName = System.Reflection.Assembly.GetAssembly(typeof(EvReceiverCalendar.EventReceiver1.EventReceiver1)).FullName;
           var title = properties.List.Title;
           if (title == "Calendario Auto")
           {
               var Title = properties.AfterProperties["Title"].ToString();
               var StartDate = properties.AfterProperties["StartDate"].ToString();
               var _EndDate = properties.AfterProperties["_EndDate"].ToString();
               var StatoAuto = properties.AfterProperties["StatoAuto"].ToString();
               var Destinazione = properties.AfterProperties["Destinazione"].ToString();
               var NrFoglio = properties.AfterProperties["NrFoglio"].ToString();
               var Matricola = properties.AfterProperties["Matricola"].ToString();
               var Data_x0020_Partenza = properties.AfterProperties["Data_x0020_Partenza"].ToString();
               var Data_x0020_Rientro = properties.AfterProperties["Data_x0020_Rientro"].ToString();
               var Commessa = properties.AfterProperties["Commessa"].ToString();

               //string connectionString = @"Password=Sas1A;Persist Security Info=True;User ID=sasa;Initial Catalog=Intranet;Data Source=.\SQLExpress";
               string connectionString = @"Password=vitrociset;Persist Security Info=True;User ID=sa;Initial Catalog=Intranet;Data Source=10.20.100.72";
               using (SqlConnection conn = new SqlConnection(connectionString))
               {
                   String insertStatement = "INSERT INTO [Intranet].[AutoAziendali].[Prenotazioni] " +
                                               " ([Title] ,[StartDate] ,[_EndDate] ,[StatoAuto] ,[Destinazione] ,[NrFoglio] ,[Matricola] ,[Data_x0020_Partenza] ,[Data_x0020_Rientro] ,[Commessa], [Evento],Origin,Lista) " +
                                               " VALUES (@Title, @StartDate, @_EndDate, @StatoAuto, @Destinazione, @NrFoglio, @Matricola, @Data_x0020_Partenza, @Data_x0020_Rientro, @Commessa, 'ItemUpdating', '" + assemblyName + "', '" + title + "')";

                   using (SqlCommand command = new SqlCommand(insertStatement, conn))
                   {
                       command.Parameters.AddWithValue("@Title", Title);
                       command.Parameters.AddWithValue("@StartDate", StartDate);
                       command.Parameters.AddWithValue("@_EndDate", _EndDate);
                       command.Parameters.AddWithValue("@StatoAuto", StatoAuto);
                       command.Parameters.AddWithValue("@Destinazione", Destinazione);
                       command.Parameters.AddWithValue("@NrFoglio", NrFoglio);
                       command.Parameters.AddWithValue("@Matricola", Matricola);
                       command.Parameters.AddWithValue("@Data_x0020_Partenza", Data_x0020_Partenza);
                       command.Parameters.AddWithValue("@Data_x0020_Rientro", Data_x0020_Rientro);
                       command.Parameters.AddWithValue("@Commessa", Commessa);

                       conn.Open();
                       int result = command.ExecuteNonQuery();

                       // Check Error
                       // if (result < 0)
                        //   Console.WriteLine("Error inserting data into Database!");
                   }
               }
           }
           base.ItemUpdating(properties);
       }

       /// <summary>
       /// An item is being deleted.
       /// </summary>
       public override void ItemDeleting(SPItemEventProperties properties)
       {
           var assemblyName = System.Reflection.Assembly.GetAssembly(typeof(EvReceiverCalendar.EventReceiver1.EventReceiver1)).FullName;
           var title = properties.List.Title;
           if (title == "Calendario Auto")
           {
               var Title = properties.ListItem["Title"] == null ? "" : properties.ListItem["Title"].ToString();
               var StartDate = properties.ListItem["StartDate"] == null ? "" : properties.ListItem["StartDate"].ToString();
               var _EndDate = properties.ListItem["_EndDate"] == null ? "" : properties.ListItem["_EndDate"].ToString();
               var StatoAuto = properties.ListItem["StatoAuto"] == null ? "" : properties.ListItem["StatoAuto"].ToString();
               var Destinazione = properties.ListItem["Destinazione"] == null ? "" : properties.ListItem["Destinazione"].ToString();
               var NrFoglio = properties.ListItem["NrFoglio"] == null ? "" : properties.ListItem["NrFoglio"].ToString();
               var Matricola = properties.ListItem["Matricola"] == null ? "" : properties.ListItem["Matricola"].ToString();
               var Data_x0020_Partenza = properties.ListItem["Data_x0020_Partenza"] == null ? "" : properties.ListItem["Data_x0020_Partenza"].ToString();
               var Data_x0020_Rientro = properties.ListItem["Data_x0020_Rientro"] == null ? "" : properties.ListItem["Data_x0020_Rientro"].ToString();
               var Commessa = properties.ListItem["Commessa"] == null ? "" : properties.ListItem["Commessa"].ToString();

               //string connectionString = @"Password=Sas1A;Persist Security Info=True;User ID=sasa;Initial Catalog=Intranet;Data Source=.\SQLExpress";
               string connectionString = @"Password=vitrociset;Persist Security Info=True;User ID=sa;Initial Catalog=Intranet;Data Source=10.20.100.72";
               using (SqlConnection conn = new SqlConnection(connectionString))
               {
                   String insertStatement = "INSERT INTO [Intranet].[AutoAziendali].[Prenotazioni] " +
                                               " ([Title] ,[StartDate] ,[_EndDate] ,[StatoAuto] ,[Destinazione] ,[NrFoglio] ,[Matricola] ,[Data_x0020_Partenza] ,[Data_x0020_Rientro] ,[Commessa], [Evento], Origin,Lista) " +
                                               " VALUES (@Title, @StartDate, @_EndDate, @StatoAuto, @Destinazione, @NrFoglio, @Matricola, @Data_x0020_Partenza, @Data_x0020_Rientro, @Commessa, 'ItemDeleting', '" + assemblyName + "', '" + title + "')";

                   using (SqlCommand command = new SqlCommand(insertStatement, conn))
                   {
                       command.Parameters.AddWithValue("@Title", Title);
                       command.Parameters.AddWithValue("@StartDate", StartDate);
                       command.Parameters.AddWithValue("@_EndDate", _EndDate);
                       command.Parameters.AddWithValue("@StatoAuto", StatoAuto);
                       command.Parameters.AddWithValue("@Destinazione", Destinazione);
                       command.Parameters.AddWithValue("@NrFoglio", NrFoglio);
                       command.Parameters.AddWithValue("@Matricola", Matricola);
                       command.Parameters.AddWithValue("@Data_x0020_Partenza", Data_x0020_Partenza);
                       command.Parameters.AddWithValue("@Data_x0020_Rientro", Data_x0020_Rientro);
                       command.Parameters.AddWithValue("@Commessa", Commessa);

                       conn.Open();
                       int result = command.ExecuteNonQuery();

                       // Check Error
                       //if (result < 0)
                       //    Console.WriteLine("Error inserting data into Database!");
                   }
               }
           }

           base.ItemDeleting(properties);
       }


    }
}
