using System;
using System.Linq;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;

namespace Generater
{
    static class ExcelLoader
    {
        #region 상수 정의
        /// <summary>
        /// 2007 이전 파일 포맷
        /// </summary>
        private const string OLD_EXTENSION = ".xls";

        /// <summary>
        /// 2007 이후 파일 포맷
        /// </summary>
        private const string NEW_EXTENSION = ".xlsx";

        /// <summary>
        /// 열거형을 구분하는 키워드
        /// </summary>
        private const string TOKEN_ENUM_TABLE = "Enum_";
        #endregion

        /// <summary>
        /// in memory 데이터 테이블을 이용하여 경로에 존재하는 엑셀 파일을 컨테이너에 특정 코드 블럭으로 저장한다.
        /// 컨테이너에 저장한 데이터는 코드 생성 및 특정 포맷으로 변경하기 위해 사용된다.
        /// </summary>
        /// <param name="excelPathWithFileName">엑셀 파일이 존재하는 경로</param>
        /// <param name="container"></param>
        public static void Initialize(string excelPath, Headergenerater container = null)
        {
            var files = Directory.GetFiles(excelPath);

            // 열거형과 데이터를 로드한다.
            for (int i = 0; i < 2; ++i)
            {
                // 테이블 정의 열거형을 먼저 로드해야 테이블헤더를 로드할 때 열거형 타입을 정상적으로 로드할 수 있다.
                foreach (var file in files)
                {
                    if (Path.GetFileName(file).Contains("~$") || Path.GetFileName(file).Contains("$"))
                    {
                        continue;
                    }

                    if (false == IsExcelExtension(Path.GetExtension(file)))
                    {
                        continue;
                    }

                    var connection = GetOleDbConnectionStringBuilder(file);

                    using (var context = new OleDbConnection(connection.ToString()))
                    {
                        context.Open();

                        using (DataTable table = context.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" }))
                        {
                            ReadSheet(context, table, container, i == 0 ? true : false);
                        }
                    }
                }
            }
        }


        /// <summary>
        /// 엑셀 확장자 여부 확인
        /// </summary>
        /// <param name="extension">파일 확장자</param>
        /// <returns>엑셀 확장자인 경우 true</returns>
        public static bool IsExcelExtension(string extension)
        {
            return OLD_EXTENSION.Equals(extension, StringComparison.OrdinalIgnoreCase) || NEW_EXTENSION.Equals(extension, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 커넥션 스트링 관리객체를 얻는다.
        /// </summary>
        /// <param name="filePath">파일 경로</param>
        /// <returns>커넥션 스트링 관리 객체</returns>
        public static OleDbConnectionStringBuilder GetOleDbConnectionStringBuilder(string filePath)
        {
            var connection = new OleDbConnectionStringBuilder
            {
                DataSource = filePath,
                Provider = Path.GetExtension(filePath).Equals(OLD_EXTENSION, StringComparison.OrdinalIgnoreCase)
                ? "Microsoft.Jet.OLEDB.4.0"
                : "Microsoft.ACE.OLEDB.12.0"
            };

            if (true == Path.GetExtension(filePath).Equals(OLD_EXTENSION, StringComparison.OrdinalIgnoreCase))
            {
                connection["Extended Properties"] = "Excel 8.0;HDR=Yes;IMEX=1";
            }
            else
            {
                connection["Extended Properties"] = "Excel 12.0;HDR=Yes;IMEX=1";
            }

            return connection;
        }


        /// <summary>
        /// 엑셀 데이터를 시트 블럭으로 읽는다.
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="table">엑셀 파일로 만들어진 in memory 데이터 테이블</param>
        /// <param name="container">데이터 블럭을 저장할 컨테이너</param>
        /// <param name="isEnumOnly">true면 열거형 타입만 읽기, false면 테이블 헤더만 읽기</param>
        private static void ReadSheet(OleDbConnection connection, DataTable table, Headergenerater container = null, bool isEnumOnly = true)
        {
            foreach (DataRow row in table.Rows)
            {
                // 하나의 시트를 읽는다.
                string sheetName = row["TABLE_NAME"].ToString();

                bool isEnumerationTypes = sheetName.Contains(TOKEN_ENUM_TABLE);

                if (true == isEnumOnly)
                {
                    // 열거형 타입만 로드
                    if (false == isEnumerationTypes)
                    {
                        continue;
                    }
                }
                else
                {
                    // 테이블 헤더만 로드
                    if (true == isEnumerationTypes)
                    {
                        continue;
                    }

                    if (true == sheetName.Contains('#'))
                    {
                        continue;
                    }
                }

                using (var excel = new DataTable { Locale = System.Globalization.CultureInfo.CurrentCulture })
                {
                    using (var adapter = new OleDbDataAdapter(string.Format("Select * from [{0}]", sheetName), connection))
                    {
                        adapter.Fill(excel);

                        if (true == isEnumerationTypes)
                        {
                            ReadEnumTable(sheetName, excel, container);
                        }
                        else
                        {
                            ReadHeaderTable(sheetName, excel, container);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 약속된 포맷으로 정의된 열거형 블럭을 읽어들인다.
        /// </summary>
        /// <param name="name">시트명</param>
        /// <param name="table">하나의 시트로 만들어진 in memory 데이터 테이블</param>
        /// <param name="container">열거형 블럭을 저장할 컨테이너</param>
        private static void ReadEnumTable(string name, DataTable table, Headergenerater container = null)
        {
            if (0 != string.Compare(name, 0, TOKEN_ENUM_TABLE, 0, TOKEN_ENUM_TABLE.Length, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string fileName = name.Remove(0, TOKEN_ENUM_TABLE.Length);

            if (null != container)
            {
                var enumTable = new EnumTable
                {
                    FileName = fileName
                };

                container.AddEnumTable(fileName, enumTable);
            }

            #region 데이터가 존재하는 첫번째 셀 컬럼명 파싱
            // offset으로 접근하기 위해 열거형 이름은 오와열을 맞춰서 정의되어 있어야한다. 
            // 먼저 종을 맞춘다.
            for (int col = 0; col < table.Columns.Count; ++col)
            {
                string tableString = table.Columns[col].ToString();

                if (true == string.IsNullOrEmpty(tableString))
                {
                    continue;
                }

                if (false == tableString.Contains(TOKEN_ENUM_TABLE))
                {
                    continue;
                }

                if (null != container)
                {
                    container.AddItemToEnumTable(fileName, tableString, -1, col);
                }
            }
            #endregion

            // 종을 맞춘 후 행을 파싱한다.
            for (var row = 0; row < table.Rows.Count; ++row)
            {
                DataRow dataRow = table.Rows[row];

                for (var col = 0; col < dataRow.ItemArray.Length; ++col)
                {
                    string rowString = dataRow.ItemArray[col].ToString();

                    if (null == container)
                    {
                        continue;
                    }

                    if (false == string.IsNullOrEmpty(rowString) && (true == rowString.Contains(TOKEN_ENUM_TABLE) || true == rowString.Contains(EnumTable.TOKEN_FLAGS)))
                    {
                        container.AddItemToEnumTable(fileName, rowString, row, col);
                    }
                    else
                    {
                        container.AddMemberToEnumTable(fileName, rowString, row, col);
                    }
                }
            }
        }


        /// <summary>
        /// 약속된 포맷으로 작성된 엑셀 데이터 테이블을 읽어들인다.
        /// </summary>
        /// <param name="name">시트명</param>
        /// <param name="table">하나의 시트로 만들어진 in memory 데이터 테이블</param>
        /// <param name="container">헤더 테이블을 저장할 컨테이너</param>
        private static void ReadHeaderTable(string name, DataTable table, Headergenerater container = null)
        {
            var tempComments = new List<string>();
            var customEnumMemberComments = new List<Tuple<string, string>>();
            CustomEnumTable tempCustomEnumTable = null;

            int indexOfID = 0;

            #region 아이디 컬럼 인덱스 체크
            DataRow memberNameRow = table.Rows[0];
            for (var col = 0; col < memberNameRow.ItemArray.Length; ++col)
            {
                var cellString = memberNameRow.ItemArray[col].ToString();
                if (cellString.Equals("ID", StringComparison.OrdinalIgnoreCase))
                {
                    indexOfID = col;
                    break;
                }
            }
            #endregion

            #region 데이터가 존재하는 첫번째 셀 컬럼명 파싱
            bool isCreate = false;
            for (int col = 0; col < table.Columns.Count; ++col)
            {
                if (col < indexOfID)
                {
                    // 아이디보다 컬럼이 많은 경우 쓰레기 데이터인 경우일 확률이 크다.
                    continue;
                }

                if (string.Format("F{0}", col+1).Equals(table.Columns[col].ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    tempComments.Add(string.Empty);
                }
                else
                {
                    tempComments.Add(table.Columns[col].ToString());
                }

                // 테이블 헤더의 첫번째 멤버는 ID이며 이는 기본적인 약속이다.
                if (true == table.Columns[col].ToString().Equals("ID", StringComparison.Ordinal))
                {
                    isCreate = true;

                    if (null != container)
                    {
                        container.CreateHeaderTable(name, -1, col, table.Columns.Count, tempComments);
                    }

                    continue;
                }

                if (true == isCreate && null != container)
                {
                    container.AddColumnToHeaderTable(name, table.Columns[col].ToString(), -1, col, table.Columns.Count);
                }

                // 테이블 이름으로 시작한다면 주석을 작성하지 않은 경우이다.
                if (true == isCreate)
                {
                    tempComments.Clear();
                }
            }
            #endregion

            if(string.Equals("CommonConfig", name.Trim('$'), StringComparison.OrdinalIgnoreCase))
            {
                tempCustomEnumTable = new CustomEnumTable();
                tempCustomEnumTable.FileName = "CommonConfig";
                tempCustomEnumTable.CommentIndex = 3;
            }

            #region 행을 파싱한다.
            for (var row = 0; row < table.Rows.Count; ++row)
            {
                DataRow dataRow = table.Rows[row];
                for (var col = 0; col < dataRow.ItemArray.Length; ++col)
                {
                    var cellString = dataRow.ItemArray[col].ToString();
                    if (true == cellString.Equals("ID", StringComparison.Ordinal))
                    {
                        if (null != container)
                        {
                            container.CreateHeaderTable(name, row, col, table.Columns.Count, tempComments);
                        }

                        isCreate = true;
                        continue;
                    }

                    if (true == isCreate && null != container)
                    {
                        container.AddColumnToHeaderTable(name, cellString, row, col, table.Columns.Count);

                        if (null != tempCustomEnumTable && 1 < row && col == tempCustomEnumTable.CommentIndex)
                        {
                            customEnumMemberComments.Add(new Tuple<string, string>(dataRow[0].ToString(), cellString));
                        }
                    }
                }
            }
            #endregion

            #region 커스텀 열거형 생성
            if (null != tempCustomEnumTable)
            {
                tempCustomEnumTable.SetData(customEnumMemberComments);
                container.AddCustomEnumTable(tempCustomEnumTable.FileName, tempCustomEnumTable);
            }
            #endregion
        }
    }
}
