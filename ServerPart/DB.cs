using RiskManagementLibrary;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServerPart
{
    class DB
    {
        public static void InsertToEventIdentifyRisk(RiskManagementLib rs)
        {
            SqlCommand command;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();

            using (command = new SqlCommand("DELETE FROM EventIdentifyRisk", connection))
            {
                command.ExecuteNonQuery();
            }

            string[] text;
            text = rs.technical;
            string sqlIns = "INSERT INTO EventIdentifyRisk (RiskName, Chance, Percents)  VALUES (@name, @information, @other)";
            try
            {
                command = new SqlCommand(sqlIns, connection);
                for (int k = 0; k < 4; k++)
                {
                    switch (k)
                    {
                        case 0: text = rs.eventTechnical; break;
                        case 1: text = rs.eventValueRisks; break;
                        case 2: text = rs.eventPlanRisks; break;
                        case 3: text = rs.eventProcesRisks; break;
                    }
                    for (int i = 0; i < rs.eventIdentifyRisk[k].Length; i++)
                    {
                        command.Parameters.AddWithValue("@name", text[i]);
                        command.Parameters.AddWithValue("@information", rs.eventIdentifyRisk[k][i]);
                        if (i == 0) command.Parameters.AddWithValue("@other", rs.eventPercentRisk[k]);
                        else { command.Parameters.AddWithValue("@other", DBNull.Value); }
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                    if (k == 3)
                    {
                        command.Parameters.AddWithValue("@name", DBNull.Value);
                        command.Parameters.AddWithValue("@information", DBNull.Value);
                        command.Parameters.AddWithValue("@other", rs.eventPercentRisk[4]);
                        command.ExecuteNonQuery();
                    }
                }
                command.Dispose();
                command = null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                connection.Close();
            }
        }

        public static void InsertToIdentifyRisk(RiskManagementLib rs)
        {
            SqlCommand command;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();

            using (command = new SqlCommand("DELETE FROM IdentifyRisk", connection))
            {
                command.ExecuteNonQuery();
            }

            string[] text;
            text = rs.technical;
            string sqlIns = "INSERT INTO IdentifyRisk (RiskName, Chance, Percents)  VALUES (@name, @information, @other)";
            try
            {
                command = new SqlCommand(sqlIns, connection);
                for (int k = 0; k < 4; k++)
                {
                    switch (k)
                    {
                        case 0: text = rs.technical; break;
                        case 1: text = rs.valueRisks; break;
                        case 2: text = rs.planRisks; break;
                        case 3: text = rs.procesRisks; break;
                    }
                    for (int i = 0; i < rs.IdentifyRisk[k].Length; i++)
                    {
                        command.Parameters.AddWithValue("@name", text[i]);
                        command.Parameters.AddWithValue("@information", rs.IdentifyRisk[k][i]);
                        if (i == 0) command.Parameters.AddWithValue("@other", rs.percentRisk[k]);
                        else { command.Parameters.AddWithValue("@other", DBNull.Value); }
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                    if (k == 3)
                    {
                        command.Parameters.AddWithValue("@name", DBNull.Value);
                        command.Parameters.AddWithValue("@information", DBNull.Value);
                        command.Parameters.AddWithValue("@other", rs.percentRisk[4]);
                        command.ExecuteNonQuery();
                    }
                }
                command.Dispose();
                command = null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                connection.Close();
            }
        }
        public static void InsertToAnalisIdentifyRisk(RiskManagementLib rs)
        {
            SqlCommand command;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();

            using (command = new SqlCommand("DELETE FROM AnalisIdentifyRisk", connection))
            {
                command.ExecuteNonQuery();
            }

            string[] text;
            text = rs.technical;
            string sqlIns = "INSERT INTO AnalisIdentifyRisk (RiskName, Expert1, Expert2, Expert3, Expert4, Expert5, Expert6, Expert7, Expert8, Expert9, Expert10, Percents)  VALUES (@name, @expert1, @expert2, @expert3, @expert4, @expert5, @expert6, @expert7, @expert8, @expert9, @expert10, @prcnt)";
            try
            {
                command = new SqlCommand(sqlIns, connection);
                for (int k = 0; k < 4; k++)
                {
                    switch (k)
                    {
                        case 0: text = rs.eventTechnical; break;
                        case 1: text = rs.eventValueRisks; break;
                        case 2: text = rs.eventPlanRisks; break;
                        case 3: text = rs.eventProcesRisks; break;
                    }
                    for (int i = 0; i < rs.analisIdentifyRisk[k].Length/11; i++)
                    {
                        command.Parameters.AddWithValue("@name", text[i]);
                        command.Parameters.AddWithValue("@expert1", rs.analisIdentifyRisk[k][i, 0]);
                        command.Parameters.AddWithValue("@expert2", rs.analisIdentifyRisk[k][i, 1]);
                        command.Parameters.AddWithValue("@expert3", rs.analisIdentifyRisk[k][i, 2]);
                        command.Parameters.AddWithValue("@expert4", rs.analisIdentifyRisk[k][i, 3]);
                        command.Parameters.AddWithValue("@expert5", rs.analisIdentifyRisk[k][i, 4]);
                        command.Parameters.AddWithValue("@expert6", rs.analisIdentifyRisk[k][i, 5]);
                        command.Parameters.AddWithValue("@expert7", rs.analisIdentifyRisk[k][i, 6]);
                        command.Parameters.AddWithValue("@expert8", rs.analisIdentifyRisk[k][i, 7]);
                        command.Parameters.AddWithValue("@expert9", rs.analisIdentifyRisk[k][i, 8]);
                        command.Parameters.AddWithValue("@expert10", rs.analisIdentifyRisk[k][i, 9]);
                        command.Parameters.AddWithValue("@prcnt", rs.analisIdentifyRisk[k][i, 10]); 
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                }
                command.Dispose();
                command = null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                connection.Close();
            }
        }
        public static void InsertToAnalisRiskWithValidity(RiskManagementLib rs)
        {
            SqlCommand command;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();

            using (command = new SqlCommand("DELETE FROM AnalisRiskWithValidity", connection))
            {
                command.ExecuteNonQuery();
            }

            string[] text;
            text = rs.technical;
            string sqlIns = "INSERT INTO AnalisRiskWithValidity (RiskName, Expert1, Expert2, Expert3, Expert4, Expert5, Expert6, Expert7, Expert8, Expert9, Expert10, Percents)  VALUES (@name, @expert1, @expert2, @expert3, @expert4, @expert5, @expert6, @expert7, @expert8, @expert9, @expert10, @prcnt)";
            try
            {
                command = new SqlCommand(sqlIns, connection);
                for (int k = 0; k < 4; k++)
                {
                    switch (k)
                    {
                        case 0: text = rs.eventTechnical; break;
                        case 1: text = rs.eventValueRisks; break;
                        case 2: text = rs.eventPlanRisks; break;
                        case 3: text = rs.eventProcesRisks; break;
                    }
                    for (int i = 0; i < rs.analisRiskWithValidity[k].Length / 11; i++)
                    {
                        command.Parameters.AddWithValue("@name", text[i]);
                        command.Parameters.AddWithValue("@expert1", rs.analisRiskWithValidity[k][i, 0]);
                        command.Parameters.AddWithValue("@expert2", rs.analisRiskWithValidity[k][i, 1]);
                        command.Parameters.AddWithValue("@expert3", rs.analisRiskWithValidity[k][i, 2]);
                        command.Parameters.AddWithValue("@expert4", rs.analisRiskWithValidity[k][i, 3]);
                        command.Parameters.AddWithValue("@expert5", rs.analisRiskWithValidity[k][i, 4]);
                        command.Parameters.AddWithValue("@expert6", rs.analisRiskWithValidity[k][i, 5]);
                        command.Parameters.AddWithValue("@expert7", rs.analisRiskWithValidity[k][i, 6]);
                        command.Parameters.AddWithValue("@expert8", rs.analisRiskWithValidity[k][i, 7]);
                        command.Parameters.AddWithValue("@expert9", rs.analisRiskWithValidity[k][i, 8]);
                        command.Parameters.AddWithValue("@expert10", rs.analisRiskWithValidity[k][i, 9]);
                        command.Parameters.AddWithValue("@prcnt", rs.analisRiskWithValidity[k][i, 10]);
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                }
                command.Dispose();
                command = null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
                connection.Close();
            }
        }
        internal static void SelectIdentifyRisk(RiskManagementLib rs)
        {
            SqlCommand command;
            int i = 0;
            int k = 0;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();
            command = new SqlCommand("Select RiskName, Chance, Percents from IdentifyRisk ", connection);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while(reader.Read())
                {
                    if(k == 0)
                    {
                        rs.technical[i] = reader.GetString(0);
                        rs.IdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.percentRisk[k] = reader.GetDouble(2);

                    }
                    if(k == 1)
                    {
                        rs.valueRisks[i] = reader.GetString(0);
                        rs.IdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.percentRisk[k] = reader.GetDouble(2);
                    }
                    if (k == 2)
                    {
                        rs.planRisks[i] = reader.GetString(0);
                        rs.IdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.percentRisk[k] = reader.GetDouble(2);
                    }
                    if (k == 3)
                    {
                        rs.procesRisks[i] = reader.GetString(0);
                        rs.IdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.percentRisk[k] = reader.GetDouble(2);
                    }
                    if (k == 4)
                    {
                        rs.percentRisk[k] = reader.GetDouble(2);
                    }
                    i++;
                    if (i == 8 && k == 0) { i = 0;k = 1; }
                    if (i == 4 && k == 1) { i = 0; k = 2; }
                    if (i == 4 && k == 2) { i = 0; k = 3; }
                    if (i == 6 && k == 3) { i = 0; k = 4; }
                }
            }
            connection.Close();
        }
        internal static void SelectEventIdentifyRisk(RiskManagementLib rs)
        {
            SqlCommand command;
            int i = 0;
            int k = 0;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();
            command = new SqlCommand("Select RiskName, Chance, Percents from EventIdentifyRisk ", connection);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (k == 0)
                    {
                        rs.eventTechnical[i] = reader.GetString(0);
                        rs.eventIdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.eventPercentRisk[k] = reader.GetDouble(2);

                    }
                    if (k == 1)
                    {
                        rs.eventValueRisks[i] = reader.GetString(0);
                        rs.eventIdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.eventPercentRisk[k] = reader.GetDouble(2);
                    }
                    if (k == 2)
                    {
                        rs.eventPlanRisks[i] = reader.GetString(0);
                        rs.eventIdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.eventPercentRisk[k] = reader.GetDouble(2);
                    }
                    if (k == 3)
                    {
                        rs.eventProcesRisks[i] = reader.GetString(0);
                        rs.eventIdentifyRisk[k][i] = reader.GetInt32(1);
                        if (i == 0) rs.eventPercentRisk[k] = reader.GetDouble(2);
                    }
                    if (k == 4)
                    {
                        rs.eventPercentRisk[k] = reader.GetDouble(2);
                    }
                    i++;
                    if (i == 12 && k == 0) { i = 0; k = 1; }
                    if (i == 10 && k == 1) { i = 0; k = 2; }
                    if (i == 12 && k == 2) { i = 0; k = 3; }
                    if (i == 17 && k == 3) { i = 0; k = 4; }
                }
            }
            connection.Close();
        }

        internal static void SelectAnalisIdentifyRisk(RiskManagementLib rs)
        {
            SqlCommand command;
            int i = 0;
            int k = 0;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();
            command = new SqlCommand("Select RiskName, Expert1, Expert2, Expert3, Expert4, Expert5, Expert6, Expert7, Expert8, Expert9, Expert10, Percents from AnalisIdentifyRisk ", connection);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (k == 0)
                    {
                        rs.eventTechnical[i] = reader.GetString(0);
                        rs.analisIdentifyRisk[k][i, 0] = reader.GetDouble(1);
                        rs.analisIdentifyRisk[k][i, 1] = reader.GetDouble(2);
                        rs.analisIdentifyRisk[k][i, 2] = reader.GetDouble(3);
                        rs.analisIdentifyRisk[k][i, 3] = reader.GetDouble(4);
                        rs.analisIdentifyRisk[k][i, 4] = reader.GetDouble(5);
                        rs.analisIdentifyRisk[k][i, 5] = reader.GetDouble(6);
                        rs.analisIdentifyRisk[k][i, 6] = reader.GetDouble(7);
                        rs.analisIdentifyRisk[k][i, 7] = reader.GetDouble(8);
                        rs.analisIdentifyRisk[k][i, 8] = reader.GetDouble(9);
                        rs.analisIdentifyRisk[k][i, 9] = reader.GetDouble(10);
                        rs.analisIdentifyRisk[k][i, 10] = reader.GetDouble(11);

                    }
                    if (k == 1)
                    {
                        rs.eventValueRisks[i] = reader.GetString(0);
                        rs.analisIdentifyRisk[k][i, 0] = reader.GetDouble(1);
                        rs.analisIdentifyRisk[k][i, 1] = reader.GetDouble(2);
                        rs.analisIdentifyRisk[k][i, 2] = reader.GetDouble(3);
                        rs.analisIdentifyRisk[k][i, 3] = reader.GetDouble(4);
                        rs.analisIdentifyRisk[k][i, 4] = reader.GetDouble(5);
                        rs.analisIdentifyRisk[k][i, 5] = reader.GetDouble(6);
                        rs.analisIdentifyRisk[k][i, 6] = reader.GetDouble(7);
                        rs.analisIdentifyRisk[k][i, 7] = reader.GetDouble(8);
                        rs.analisIdentifyRisk[k][i, 8] = reader.GetDouble(9);
                        rs.analisIdentifyRisk[k][i, 9] = reader.GetDouble(10);
                        rs.analisIdentifyRisk[k][i, 10] = reader.GetDouble(11);
                    }
                    if (k == 2)
                    {
                        rs.eventPlanRisks[i] = reader.GetString(0);
                        rs.analisIdentifyRisk[k][i, 0] = reader.GetDouble(1);
                        rs.analisIdentifyRisk[k][i, 1] = reader.GetDouble(2);
                        rs.analisIdentifyRisk[k][i, 2] = reader.GetDouble(3);
                        rs.analisIdentifyRisk[k][i, 3] = reader.GetDouble(4);
                        rs.analisIdentifyRisk[k][i, 4] = reader.GetDouble(5);
                        rs.analisIdentifyRisk[k][i, 5] = reader.GetDouble(6);
                        rs.analisIdentifyRisk[k][i, 6] = reader.GetDouble(7);
                        rs.analisIdentifyRisk[k][i, 7] = reader.GetDouble(8);
                        rs.analisIdentifyRisk[k][i, 8] = reader.GetDouble(9);
                        rs.analisIdentifyRisk[k][i, 9] = reader.GetDouble(10);
                        rs.analisIdentifyRisk[k][i, 10] = reader.GetDouble(11);
                    }
                    if (k == 3)
                    {
                        rs.eventProcesRisks[i] = reader.GetString(0);
                        rs.analisIdentifyRisk[k][i, 0] = reader.GetDouble(1);
                        rs.analisIdentifyRisk[k][i, 1] = reader.GetDouble(2);
                        rs.analisIdentifyRisk[k][i, 2] = reader.GetDouble(3);
                        rs.analisIdentifyRisk[k][i, 3] = reader.GetDouble(4);
                        rs.analisIdentifyRisk[k][i, 4] = reader.GetDouble(5);
                        rs.analisIdentifyRisk[k][i, 5] = reader.GetDouble(6);
                        rs.analisIdentifyRisk[k][i, 6] = reader.GetDouble(7);
                        rs.analisIdentifyRisk[k][i, 7] = reader.GetDouble(8);
                        rs.analisIdentifyRisk[k][i, 8] = reader.GetDouble(9);
                        rs.analisIdentifyRisk[k][i, 9] = reader.GetDouble(10);
                        rs.analisIdentifyRisk[k][i, 10] = reader.GetDouble(11);
                    }
                    i++;
                    if (i == 12 && k == 0) { i = 0; k = 1; }
                    if (i == 10 && k == 1) { i = 0; k = 2; }
                    if (i == 12 && k == 2) { i = 0; k = 3; }
                    if (i == 17 && k == 3) { i = 0; k = 4; }
                }
            }
            connection.Close();
        }
        internal static void SelectAnalisRiskWithValidity(RiskManagementLib rs)
        {
            SqlCommand command;
            int i = 0;
            int k = 0;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();
            command = new SqlCommand("Select RiskName, Expert1, Expert2, Expert3, Expert4, Expert5, Expert6, Expert7, Expert8, Expert9, Expert10, Percents from AnalisRiskWithValidity ", connection);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (k == 0)
                    {
                        rs.eventTechnical[i] = reader.GetString(0);
                        rs.analisRiskWithValidity[k][i, 0] = reader.GetDouble(1);
                        rs.analisRiskWithValidity[k][i, 1] = reader.GetDouble(2);
                        rs.analisRiskWithValidity[k][i, 2] = reader.GetDouble(3);
                        rs.analisRiskWithValidity[k][i, 3] = reader.GetDouble(4);
                        rs.analisRiskWithValidity[k][i, 4] = reader.GetDouble(5);
                        rs.analisRiskWithValidity[k][i, 5] = reader.GetDouble(6);
                        rs.analisRiskWithValidity[k][i, 6] = reader.GetDouble(7);
                        rs.analisRiskWithValidity[k][i, 7] = reader.GetDouble(8);
                        rs.analisRiskWithValidity[k][i, 8] = reader.GetDouble(9);
                        rs.analisRiskWithValidity[k][i, 9] = reader.GetDouble(10);
                        rs.analisRiskWithValidity[k][i, 10] = reader.GetDouble(11);

                    }
                    if (k == 1)
                    {
                        rs.eventValueRisks[i] = reader.GetString(0);
                        rs.analisRiskWithValidity[k][i, 0] = reader.GetDouble(1);
                        rs.analisRiskWithValidity[k][i, 1] = reader.GetDouble(2);
                        rs.analisRiskWithValidity[k][i, 2] = reader.GetDouble(3);
                        rs.analisRiskWithValidity[k][i, 3] = reader.GetDouble(4);
                        rs.analisRiskWithValidity[k][i, 4] = reader.GetDouble(5);
                        rs.analisRiskWithValidity[k][i, 5] = reader.GetDouble(6);
                        rs.analisRiskWithValidity[k][i, 6] = reader.GetDouble(7);
                        rs.analisRiskWithValidity[k][i, 7] = reader.GetDouble(8);
                        rs.analisRiskWithValidity[k][i, 8] = reader.GetDouble(9);
                        rs.analisRiskWithValidity[k][i, 9] = reader.GetDouble(10);
                        rs.analisRiskWithValidity[k][i, 10] = reader.GetDouble(11);
                    }
                    if (k == 2)
                    {
                        rs.eventPlanRisks[i] = reader.GetString(0);
                        rs.analisRiskWithValidity[k][i, 0] = reader.GetDouble(1);
                        rs.analisRiskWithValidity[k][i, 1] = reader.GetDouble(2);
                        rs.analisRiskWithValidity[k][i, 2] = reader.GetDouble(3);
                        rs.analisRiskWithValidity[k][i, 3] = reader.GetDouble(4);
                        rs.analisRiskWithValidity[k][i, 4] = reader.GetDouble(5);
                        rs.analisRiskWithValidity[k][i, 5] = reader.GetDouble(6);
                        rs.analisRiskWithValidity[k][i, 6] = reader.GetDouble(7);
                        rs.analisRiskWithValidity[k][i, 7] = reader.GetDouble(8);
                        rs.analisRiskWithValidity[k][i, 8] = reader.GetDouble(9);
                        rs.analisRiskWithValidity[k][i, 9] = reader.GetDouble(10);
                        rs.analisRiskWithValidity[k][i, 10] = reader.GetDouble(11);
                    }
                    if (k == 3)
                    {
                        rs.eventProcesRisks[i] = reader.GetString(0);
                        rs.analisRiskWithValidity[k][i, 0] = reader.GetDouble(1);
                        rs.analisRiskWithValidity[k][i, 1] = reader.GetDouble(2);
                        rs.analisRiskWithValidity[k][i, 2] = reader.GetDouble(3);
                        rs.analisRiskWithValidity[k][i, 3] = reader.GetDouble(4);
                        rs.analisRiskWithValidity[k][i, 4] = reader.GetDouble(5);
                        rs.analisRiskWithValidity[k][i, 5] = reader.GetDouble(6);
                        rs.analisRiskWithValidity[k][i, 6] = reader.GetDouble(7);
                        rs.analisRiskWithValidity[k][i, 7] = reader.GetDouble(8);
                        rs.analisRiskWithValidity[k][i, 8] = reader.GetDouble(9);
                        rs.analisRiskWithValidity[k][i, 9] = reader.GetDouble(10);
                        rs.analisRiskWithValidity[k][i, 10] = reader.GetDouble(11);
                    }
                    i++;
                    if (i == 12 && k == 0) { i = 0; k = 1; }
                    if (i == 10 && k == 1) { i = 0; k = 2; }
                    if (i == 12 && k == 2) { i = 0; k = 3; }
                    if (i == 17 && k == 3) { i = 0; k = 4; }
                }
            }
            connection.Close();
        }
        internal static bool SelectUsers(RiskManagementLib rs)
        {
            SqlCommand command;
            SqlConnection connection = new SqlConnection("Server=KALYAN\\MSSQLSERVER01; Database=RiskManagementDB; Integrated Security=true;");
            connection.Open();
            command = new SqlCommand("Select LoginUser, ParolUser, number from Users ", connection);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (rs.Login == reader.GetString(0) && rs.Parol == reader.GetString(1))
                    {
                        rs.number = reader.GetInt32(2);
                        return true;
                    }
                }
                return false;
            }
            connection.Close();
        }
    }
}
