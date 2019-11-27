using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MMCSirUtilities
{
    public class ParameterWrapper
    {
        private SqlDbType _Type;
        private object _Value;
        private string _Name;
        private ParameterDirection _Direction = ParameterDirection.Input;
        private string _TypeName = string.Empty;

        public string TypeName {
            get {
                return _TypeName;
            }
            set {
                _TypeName = value;
            }
        }

        public ParameterDirection Direction {
            get {
                return _Direction;
            }
            set {
                _Direction = value;
            }
        }
        public SqlDbType Type {
            get {
                return _Type;
            }
            set {
                _Type = value;
            }
        }
        public string Name {
            get {
                return _Name;
            }
            set {
                _Name = value;
            }
        }
        public object Value {
            get {
                return _Value;
            }
            set {
                _Value = value;
            }
        }


        public ParameterWrapper( string name, SqlDbType type, object value )
        {
            this._Type = type;
            this._Name = name;
            this._Value = value;
        }

        public ParameterWrapper( string name, SqlDbType type, object value, System.Data.ParameterDirection direction )
        {
            this._Type = type;
            this._Name = name;
            this._Value = value;
            this._Direction = direction;
        }

        public ParameterWrapper( string name, SqlDbType type, object value, System.Data.ParameterDirection direction, string typename )
        {
            this._Type = type;
            this._Name = name;
            this._Value = value;
            this._Direction = direction;
            this._TypeName = typename;
        }

        public SqlParameter GetSQLParameter()
        {
            SqlParameter retValue = new SqlParameter();
            retValue.ParameterName = this._Name;
            retValue.SqlDbType = _Type;
            retValue.Direction = this._Direction;
            retValue.Value = this._Value;
            retValue.TypeName = this._TypeName;
            return retValue;
        }
    }

    public class ParameterWrapperCollection : CollectionBase, IEnumerable<ParameterWrapper>
    {

        /// <summary>
        /// Add a new element to the columns collection
        /// </summary>
        /// <param name="objColumn">New column to Add</param>
        public void Add( ParameterWrapper ObjColumn )
        {
            try
            {

                if ( !( ObjColumn is ParameterWrapper ) )
                {
                    throw new ArgumentException("ParameterWrapper Collection can only contains ParameterWrapper Controls");
                }
                List.Add(ObjColumn);
            } catch ( Exception ex )
            {
                throw ex;
            }

        }

        /// <summary>
        /// Add a new element to the columns collection
        /// </summary>
        /// <param name="id">Id of the new column</param>
        /// <param name="label">_Label for the new column</param>
        public void Add( string name, System.Data.SqlDbType type, object value )
        {
            try
            {
                ParameterWrapper TmpItem = new ParameterWrapper(name, type, value);
                List.Add(TmpItem);
                TmpItem = null;
            } catch ( Exception ex )
            {
                throw ex;
            }

        }

        public void Add( string name, System.Data.SqlDbType type, object value, System.Data.ParameterDirection direction )
        {
            try
            {
                ParameterWrapper TmpItem = new ParameterWrapper(name, type, value, direction);
                List.Add(TmpItem);
                TmpItem = null;
            } catch ( Exception ex )
            {
                throw ex;
            }

        }

        public void Add( string name, System.Data.SqlDbType type, object value, System.Data.ParameterDirection direction, string typename )
        {
            try
            {
                ParameterWrapper TmpItem = new ParameterWrapper(name, type, value, direction, typename);
                List.Add(TmpItem);
                TmpItem = null;
            } catch ( Exception ex )
            {
                throw ex;
            }

        }

        /// <summary>
        /// Remove an element from the collection
        /// </summary>
        /// <param name="index">Position to remove in the collection</param>
        public void Remove( int Index )
        {
            if ( Index > Count - 1 || Index < 0 )
            {
                throw ( new Exception("No Item in supplied index") );
            } else
            {
                List.RemoveAt(Index);
            }
        }

        /// <summary>
        /// Remove an element from the collection
        /// </summary>
        /// <param name="id">Id of the column to remove</param>
        public void Remove( string name )
        {
            int index = -1;
            if ( this.Count > 0 )
            {
                for ( int i = 0; i < this.Count; i++ )
                {
                    if ( ( (ParameterWrapper)this[i] ).Name == name )
                    {
                        index = i;
                        break;
                    }
                }
                if ( index != -1 )
                {
                    List.RemoveAt(index);
                }
            }
        }

        /// <summary>
        /// Gets an element from the collection
        /// </summary>
        /// <param name="Index">Position in the collection</param>
        /// <returns>Column in the given position</returns>
        public ParameterWrapper this[int Index] {
            get {
                return (ParameterWrapper)this[Index];
            }
            set {
                if ( value != null )
                {
                    this[Index] = value;
                }
            }
        }

        /// <summary>
        /// Gets an element from the collection
        /// </summary>
        /// <param name="Id">Column Id</param>
        /// <returns>Column with the given Id</returns>
        public ParameterWrapper this[string Id] {
            get {
                ParameterWrapper ObjretItem = null;
                foreach ( ParameterWrapper thisCol in this )
                {
                    if ( thisCol.Name == Id )
                    {
                        ObjretItem = thisCol;
                        break;
                    }
                }
                return ObjretItem;
            }
            set {
                if ( value != null )
                {
                    int SelIndex = -1;
                    for ( int i = 0; i < Count; i++ )
                    {
                        ParameterWrapper ThisCol = (ParameterWrapper)this[i];
                        if ( ThisCol.Name == Id )
                        {
                            SelIndex = i;
                            break;
                        }
                    }
                    if ( SelIndex >= 0 )
                    {
                        this[SelIndex] = value;
                    }


                }
            }
        }

        public new IEnumerator<ParameterWrapper> GetEnumerator()
        {
            foreach ( ParameterWrapper CurrentColumnValue in this.List )
            {
                yield return CurrentColumnValue;
            }
        }

    }
}
