using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MMCSirUtilities {
    public class MMCAdapter {
        public static string MMConectionString = string.Empty;
    }

    public class SQLConnWrapper : MMCAdapter, IDisposable {

        #region Implementing Dispose
        private IntPtr handle;
        private Component component = new Component();
        private bool disposed = false;
        public SQLConnWrapper(IntPtr handle) {
            this.handle = handle;
        }
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing) {
            if (!this.disposed) {
                if (disposing) {
                    component.Dispose();
                }
                CloseHandle(handle);
                handle = IntPtr.Zero;
            }
            disposed = true;
        }
        [System.Runtime.InteropServices.DllImport("Kernel32")]
        private extern static Boolean CloseHandle(IntPtr handle);
        ~SQLConnWrapper() {
            Dispose(false);
        }
        #endregion

        private string _CurrentConnectionString = "";
        private ParameterWrapperCollection _Parameters = new ParameterWrapperCollection();

        public ParameterWrapperCollection Parameters {
            get {
                return _Parameters;
            }
            set {
                _Parameters = value;
            }
        }

        public SQLConnWrapper() {
            _CurrentConnectionString = MMConectionString;
            if (_CurrentConnectionString == string.Empty) {
                throw new Exception("No Connection string for current environment.");
            }
        }

        public SQLConnWrapper(string customConnetionString) {
            _CurrentConnectionString = customConnetionString;
            if (_CurrentConnectionString == string.Empty) {
                throw new Exception("No Connection string for current environment.");
            }
        }

        public DataTable ExecuteQuery(string procedure, int timeout = 0) {
            DataTable retValue = null;
            SqlConnection sqlConn = null;
            SqlCommand sqlComm = null;
            SqlDataReader sqlReader = null;
            using (sqlConn = new SqlConnection(_CurrentConnectionString)) {
                sqlConn.Open();
                if (sqlConn.State == ConnectionState.Open) {
                    sqlComm = new SqlCommand(procedure, sqlConn);
                    sqlComm.CommandTimeout = timeout;
                    sqlComm.CommandType = CommandType.Text;
                    sqlReader = sqlComm.ExecuteReader();
                    retValue = new DataTable();
                    retValue.Load(sqlReader);
                    sqlReader.Close();
                    foreach (SqlParameter thisParam in sqlComm.Parameters) {
                        if (thisParam.Direction == ParameterDirection.Output) {
                            this._Parameters[thisParam.ParameterName].Value = thisParam.Value;
                        }
                    }
                    sqlConn.Close();
                } else {
                    throw new Exception("Connection not open");
                }
            }
            sqlConn.Dispose();
            sqlComm.Dispose();
            sqlReader.Dispose();
            return retValue;
        }

        public DataTable ExecuteStoredProcedureToDataTable(string procedure, int timeout = 0) {
            DataTable retValue = null;
            SqlConnection sqlConn = null;
            SqlCommand sqlComm = null;
            SqlDataReader sqlReader = null;
            using (sqlConn = new SqlConnection(_CurrentConnectionString)) {
                sqlConn.Open();
                if (sqlConn.State == ConnectionState.Open) {
                    sqlComm = new SqlCommand(procedure, sqlConn);
                    sqlComm.CommandTimeout = timeout;
                    sqlComm.CommandType = CommandType.StoredProcedure;
                    if (this._Parameters.Count > 0) {
                        foreach (ParameterWrapper thisParam in this._Parameters) {
                            sqlComm.Parameters.Add(thisParam.GetSQLParameter());
                        }
                    }
                    sqlReader = sqlComm.ExecuteReader();
                    retValue = new DataTable();
                    retValue.Load(sqlReader);
                    sqlReader.Close();
                    foreach (SqlParameter thisParam in sqlComm.Parameters) {
                        if (thisParam.Direction == ParameterDirection.Output) {
                            this._Parameters[thisParam.ParameterName].Value = thisParam.Value;
                        }
                    }
                    sqlConn.Close();

                } else {
                    throw new Exception("Connection not open");
                }
            }
            sqlConn.Dispose();
            sqlComm.Dispose();
            sqlReader.Dispose();
            return retValue;
        }

        public DataTable ExecuteQueryToDataTable(string query, int timeout = 30) {
            DataTable retValue = null;
            SqlConnection sqlConn = null;
            SqlCommand sqlComm = null;
            SqlDataReader sqlReader = null;

            using (sqlConn = new SqlConnection(_CurrentConnectionString)) {
                sqlConn.Open();
                if (sqlConn.State == ConnectionState.Open) {
                    sqlComm = new SqlCommand(query, sqlConn);
                    sqlComm.CommandTimeout = timeout;
                    sqlComm.CommandType = CommandType.Text;
                    if (this._Parameters.Count > 0) {
                        foreach (ParameterWrapper thisParam in this._Parameters) {
                            sqlComm.Parameters.Add(thisParam.GetSQLParameter());
                        }
                    }
                    sqlReader = sqlComm.ExecuteReader();
                    retValue = new DataTable();
                    retValue.Load(sqlReader);
                    sqlReader.Close();
                    foreach (SqlParameter thisParam in sqlComm.Parameters) {
                        if (thisParam.Direction == ParameterDirection.Output) {
                            this._Parameters[thisParam.ParameterName].Value = thisParam.Value;
                        }
                    }
                    sqlConn.Close();

                } else {
                    throw new Exception("Connection not open");
                }
            }
            sqlConn.Dispose();
            sqlComm.Dispose();
            sqlReader.Dispose();
            return retValue;
        }

        public Boolean ExecuteStoredProcedure(string procedure, int timeout = 30) {
            Boolean retValue = true;
            SqlConnection sqlConn = null;
            SqlCommand sqlComm = null;

            using (sqlConn = new SqlConnection(_CurrentConnectionString)) {
                sqlConn.Open();
                sqlComm = new SqlCommand(procedure, sqlConn);
                sqlComm.CommandTimeout = timeout;
                sqlComm.CommandType = CommandType.StoredProcedure;
                if (this._Parameters.Count > 0) {
                    foreach (ParameterWrapper thisParam in this._Parameters) {
                        sqlComm.Parameters.Add(thisParam.GetSQLParameter());
                    }
                }
                sqlComm.ExecuteNonQuery();
                var Outputs = from p in this._Parameters
                              where p.Direction == ParameterDirection.Output
                              select p;

                foreach (ParameterWrapper thisParam in Outputs) {
                    thisParam.Value = sqlComm.Parameters[thisParam.Name].Value;
                }

                retValue = true;
                sqlConn.Close();
            }
            sqlConn.Dispose();
            sqlComm.Dispose();
            return retValue;
        }

        public Boolean ExecuteStoredProcsInTransaction(List<SQLSPWrapper> procsToExecute, int timeout = 30) {
            Boolean retValue = true;
            SqlConnection sqlConn = null;
            SqlTransaction sqlTransact = null;

            try {
                sqlConn = new SqlConnection(_CurrentConnectionString);

                sqlConn.Open();
                sqlTransact = sqlConn.BeginTransaction();

                foreach (SQLSPWrapper currentSP in procsToExecute) {
                    SqlCommand sqlComm = null;
                    sqlComm = new SqlCommand(currentSP.SPName, sqlConn, sqlTransact);
                    sqlComm.CommandTimeout = timeout;
                    sqlComm.CommandType = CommandType.StoredProcedure;
                    if (currentSP.Parameters.Count > 0) {
                        foreach (ParameterWrapper thisParam in currentSP.Parameters) {
                            sqlComm.Parameters.Add(thisParam.GetSQLParameter());
                        }
                    }
                    sqlComm.ExecuteNonQuery();
                }
                sqlTransact.Commit();

            } catch (SqlException) {
                sqlTransact.Rollback();
                retValue = false;
                throw;
            }

            sqlConn.Close();
            sqlConn.Dispose();
            return retValue;
        }

        /// <summary>
        /// This function return a dataset with the results gathered from the stored procedure executed.
        /// </summary>
        /// <param name="procedure">Name of the stored procedure to execute.</param>
        /// <returns></returns>
        public DataSet ExecuteSPToDataSet(string procedure, int timeout = 30) {
            DataSet retValue = null;
            SqlConnection sqlConn = null;
            SqlCommand sqlComm = null;
            SqlDataAdapter sqlDataAdapter = null;
            using (sqlConn = new SqlConnection(_CurrentConnectionString)) {
                sqlConn.Open();
                if (sqlConn.State == ConnectionState.Open) {
                    sqlComm = new SqlCommand(procedure, sqlConn);
                    sqlComm.CommandTimeout = timeout;
                    sqlComm.CommandType = CommandType.StoredProcedure;
                    if (this._Parameters.Count > 0) {
                        foreach (ParameterWrapper thisParam in this._Parameters) {
                            sqlComm.Parameters.Add(thisParam.GetSQLParameter());
                        }
                    }
                    sqlDataAdapter = new SqlDataAdapter(sqlComm);
                    retValue = new DataSet();
                    sqlDataAdapter.Fill(retValue);
                    sqlDataAdapter.Dispose();
                    foreach (SqlParameter thisParam in sqlComm.Parameters) {
                        if (thisParam.Direction == ParameterDirection.Output) {
                            this._Parameters[thisParam.ParameterName].Value = thisParam.Value;
                        }
                    }
                    sqlConn.Close();
                } else {
                    throw new Exception("Connection not open");
                }
            }
            sqlConn.Dispose();
            sqlComm.Dispose();
            sqlDataAdapter.Dispose();
            return retValue;
        }

    }

    public class SQLSPWrapper {
        private string _SPName = string.Empty;
        private ParameterWrapperCollection _Parameters = new ParameterWrapperCollection();
        public ParameterWrapperCollection Parameters {
            get {
                return _Parameters;
            }
            set {
                _Parameters = value;
            }
        }
        public string SPName {
            get {
                return _SPName;
            }
            set {
                _SPName = value;
            }
        }
        public SQLSPWrapper(string spname) {
            _SPName = spname;
        }
    }
}
