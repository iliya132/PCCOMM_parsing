using System;
using System.Data.SqlClient;

namespace OFKO_Robot
{
    /// <summary>
    /// Класс реализующий сбор статистики об использовании робота для дальнейшего анализа
    /// </summary>
    public class Statistic
    {
        int Robot_id = 0;
        int operationsDone = 0;
        DateTime dateTime = DateTime.Now;

        public Statistic(int robot_id)
        {
            Robot_id = robot_id;
        }

        /// <summary>
        /// Посчитать завершенную операцию
        /// </summary>
        public void OperationDone()
        {
            operationsDone++;
        }

        /// <summary>
        /// Передать статистику на сервер
        /// </summary>
        public void Commit()
        {
            SqlConnectionStringBuilder sqlConnection = new SqlConnectionStringBuilder();
            sqlConnection.InitialCatalog = "RPA_statistic";
            sqlConnection.DataSource = @"A105512\A105512";
            sqlConnection.IntegratedSecurity = true;

            using (SqlConnection connection = new SqlConnection(sqlConnection.ConnectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(@"insert into Automation (RobotId, OperationsCount, RunDate)
                                                        values(@robotId, @operationCount, @RunDate)", connection);
                command.Parameters.AddWithValue("@robotId", Robot_id);
                command.Parameters.AddWithValue("@operationCount", operationsDone);
                command.Parameters.AddWithValue("@RunDate", dateTime);
                command.ExecuteNonQuery();
            }
        }
    }
}
