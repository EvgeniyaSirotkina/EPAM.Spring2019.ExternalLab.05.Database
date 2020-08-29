using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;

namespace Database
{
    public class Program
    {
        const string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""..\..\db\db.mdb""";

        public static void Main(string[] args)
        {
            var con = new OleDbConnection(connectionString);
            var queue = new Queue<int>();
            OleDbCommand command = null;
            OleDbDataReader reader = null;

            try
            {
                // 1 Removing records from the Frequencies table.
                con.Open();
                command = con.CreateCommand();
                try
                {
                    command.CommandText = "SELECT int(ABS(x1 - x2) + 0.5) AS len, Count(*) AS num " +
                                            "FROM Coordinates " +
                                            "GROUP BY int(ABS(x1 - x2) + 0.5) " +
                                            "ORDER BY int(ABS(x1 - x2) + 0.5) ASC";
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        queue.Enqueue((int)reader.GetDouble(0));
                        queue.Enqueue(reader.GetInt32(1));
                    }
                }
                catch (IndexOutOfRangeException e)
                {
                    var stackTrace = new StackTrace(e, true);
                    var stackFrame = stackTrace.GetFrame(stackTrace.FrameCount - 1);

                    throw new IndexOutOfRangeException(ErrorConstants.IndexOutOfTheRange + stackFrame.GetFileLineNumber());
                }
                catch (OleDbException e)
                {
                    var stackTrace = new StackTrace(e, true);
                    var stackFrame = stackTrace.GetFrame(stackTrace.FrameCount - 1);

                    throw new IndexOutOfRangeException(ErrorConstants.IncorrectSqlquery + stackFrame.GetFileLineNumber());
                }
                finally
                {
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }

                // 2 Adding a record to Frequencies table.
                try
                {
                    command.CommandText = "DELETE * FROM Frequencies";
                    command.ExecuteNonQuery();

                    command.CommandText = "INSERT INTO Frequencies(len, num) VALUES(@Len, @Num)";

                    command.Parameters.Add(new OleDbParameter("@Len", OleDbType.Integer));
                    command.Parameters.Add(new OleDbParameter("@Num", OleDbType.Integer));

                    while (queue.Count != 0)
                    {
                        command.Parameters["@Len"].Value = queue.Dequeue();
                        command.Parameters["@Num"].Value = queue.Dequeue();
                        command.ExecuteNonQuery();
                    }

                    command.CommandText = "SELECT * FROM Frequencies";
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        Console.WriteLine(reader["len"] + ";" + reader["num"]);
                    }
                }
                catch (IndexOutOfRangeException e)
                {
                    var stackTrace = new StackTrace(e, true);
                    var stackFrame = stackTrace.GetFrame(stackTrace.FrameCount - 1);

                    throw new IndexOutOfRangeException(ErrorConstants.IndexOutOfTheRange + stackFrame.GetFileLineNumber());
                }
                catch (OleDbException e)
                {
                    var stackTrace = new StackTrace(e, true);
                    var stackFrame = stackTrace.GetFrame(stackTrace.FrameCount - 1);

                    throw new IndexOutOfRangeException(ErrorConstants.IncorrectSqlquery + stackFrame.GetFileLineNumber());
                }
                finally
                {
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }

                // 3 Selecting records from the Frequencies table.
                try
                {
                    command.CommandText = "SELECT len, num FROM Frequencies WHERE len > num";
                    reader = command.ExecuteReader();

                    Console.WriteLine();
                    while (reader.Read())
                    {
                        Console.WriteLine(reader["len"] + ";" + reader["num"]);
                    }
                }
                catch (IndexOutOfRangeException e)
                {
                    var stackTrace = new StackTrace(e, true);
                    var stackFrame = stackTrace.GetFrame(stackTrace.FrameCount - 1);

                    throw new IndexOutOfRangeException(ErrorConstants.IndexOutOfTheRange + stackFrame.GetFileLineNumber());
                }
                catch (OleDbException e)
                {
                    var stackTrace = new StackTrace(e, true);
                    var stackFrame = stackTrace.GetFrame(stackTrace.FrameCount - 1);

                    throw new IndexOutOfRangeException(ErrorConstants.IncorrectSqlquery + stackFrame.GetFileLineNumber());
                }
                finally
                {
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
            }
            catch (OleDbException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (IndexOutOfRangeException e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                con.Close();
            }

            Console.ReadKey();
        }
    }
}
