using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace Avantik.Web.Service.Helpers.Database
{
    public class DBHelper : IDisposable
    {
        string _ConnectionString;
        SqlConnection _conn;
        SqlTransaction _trans = null;

        public DBHelper(string connectionString)
        {
            _ConnectionString = connectionString;
        }

        public SqlCommand CreateCommand(string qry, CommandType type, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    SqlCommand cmd = new SqlCommand(qry, _conn);

                    // Associate with current transaction, if any
                    if (_trans != null)
                        cmd.Transaction = _trans;

                    // Set command type
                    cmd.CommandType = type;

                    // Construct SQL parameters
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (args[i] is string && i < (args.Length - 1))
                        {
                            SqlParameter parm = new SqlParameter();
                            parm.ParameterName = (string)args[i];
                            parm.Value = args[++i];
                            cmd.Parameters.Add(parm);
                        }
                        else if (args[i] is SqlParameter)
                        {
                            cmd.Parameters.Add((SqlParameter)args[i]);
                        }
                        else throw new ArgumentException("Invalid number or type of arguments supplied");
                    }
                    return cmd;
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        #region Exec Members

        /// <summary>
        /// Executes a query that returns no results
        /// </summary>
        /// <param name="qry">Query text</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>The number of rows affected</returns>
        public int ExecNonQuery(string qry, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    using (SqlCommand cmd = CreateCommand(qry, CommandType.Text, args))
                    {
                        return cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Executes a stored procedure that returns no results
        /// </summary>
        /// <param name="proc">Name of stored proceduret</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>The number of rows affected</returns>
        public int ExecNonQueryProc(string proc, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(proc) == false)
                {
                    using (SqlCommand cmd = CreateCommand(proc, CommandType.StoredProcedure, args))
                    {
                        return cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Store procedure name is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Executes a query that returns a single value
        /// </summary>
        /// <param name="qry">Query text</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>Value of first column and first row of the results</returns>
        public object ExecScalar(string qry, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    using (SqlCommand cmd = CreateCommand(qry, CommandType.Text, args))
                    {
                        return cmd.ExecuteScalar();
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Executes a query that returns a single value
        /// </summary>
        /// <param name="proc">Name of stored proceduret</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>Value of first column and first row of the results</returns>
        public object ExecScalarProc(string qry, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    using (SqlCommand cmd = CreateCommand(qry, CommandType.StoredProcedure, args))
                    {
                        return cmd.ExecuteScalar();
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            
        }

        /// <summary>
        /// Executes a query and returns the results as a SqlDataReader
        /// </summary>
        /// <param name="qry">Query text</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>Results as a SqlDataReader</returns>
        public SqlDataReader ExecDataReader(string qry, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    using (SqlCommand cmd = CreateCommand(qry, CommandType.Text, args))
                    {
                        return cmd.ExecuteReader();
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            
        }

        /// <summary>
        /// Executes a stored procedure and returns the results as a SqlDataReader
        /// </summary>
        /// <param name="proc">Name of stored proceduret</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>Results as a SqlDataReader</returns>
        public SqlDataReader ExecDataReaderProc(string proc, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(proc) == false)
                {
                    using (SqlCommand cmd = CreateCommand(proc, CommandType.StoredProcedure, args))
                    {
                        return cmd.ExecuteReader();
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Store procedure name is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            
        }

        /// <summary>
        /// Executes a query and returns the results as a DataSet
        /// </summary>
        /// <param name="qry">Query text</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>Results as a DataSet</returns>
        public DataSet ExecDataSet(string qry, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    using (SqlCommand cmd = CreateCommand(qry, CommandType.Text, args))
                    {
                        SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                        DataSet ds = new DataSet();
                        adapt.Fill(ds);
                        return ds;
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Executes a stored procedure and returns the results as a Data Set
        /// </summary>
        /// <param name="proc">Name of stored proceduret</param>
        /// <param name="args">Any number of parameter name/value pairs and/or SQLParameter arguments</param>
        /// <returns>Results as a DataSet</returns>
        public DataSet ExecDataSetProc(string qry, params object[] args)
        {
            try
            {
                if (string.IsNullOrEmpty(qry) == false)
                {
                    using (SqlCommand cmd = CreateCommand(qry, CommandType.StoredProcedure, args))
                    {
                        SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                        DataSet ds = new DataSet();
                        adapt.Fill(ds);
                        return ds;
                    }
                }
                else
                {
                    throw new ArgumentException("DBHelper: Query parameter is required.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        #endregion

        #region Transaction Members
        /// <summary>
        /// Begins a transaction
        /// </summary>
        /// <returns>The new SqlTransaction object</returns>
        public void BeginTransaction()
        {
            try
            {
                Rollback();
                _trans = _conn.BeginTransaction();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Commits any transaction in effect.
        /// </summary>
        public void Commit()
        {
            try
            {
                if (_trans != null)
                {
                    _trans.Commit();
                    _trans = null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            
        }

        /// <summary>
        /// Rolls back any transaction in effect.
        /// </summary>
        public void Rollback()
        {
            try
            {
                if (_trans != null)
                {
                    _trans.Rollback();
                    _trans = null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        #endregion

        
        #region Helper
        public void Connect()
        {
            try
            {
                if (string.IsNullOrEmpty(_ConnectionString) == false)
                {
                    _conn = new SqlConnection(_ConnectionString);
                    _conn.Open();
                }
                else
                {
                    throw new ArgumentException("DBHelper: Required connection string.");
                }
              
            }
            catch(Exception ex)
            {
                throw;
            }
            
        }

        private void CloseConnection()
        {
            try
            {
                if (_conn != null)
                {
                    _conn.Close();
                    _conn.Dispose();
                    _conn = null;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            
        }
        #endregion
        public void Dispose()
        {
            CloseConnection();
        }
    }
}
