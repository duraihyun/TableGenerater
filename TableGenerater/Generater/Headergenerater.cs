using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace Generater
{
    /// <summary>
    /// 파일 헤더 생성
    /// </summary>
    public class Headergenerater
    {
        private readonly List<string> headerNames = new List<string>();

        /// <summary>
        /// 열거형 테이블 컨테이너, 해당 정보를 바탕으로 코드로 가공한다.
        /// 키: 테이블 명
        /// 값: 엑셀 시트에서 로드한 열거형 정보
        /// </summary>
        private readonly Dictionary<string, EnumTable> enumTables = new Dictionary<string, EnumTable>();

        /// <summary>
        /// 테이블 헤더 컨테이너, 해당 정보를 바탕으로 코드나 특정 포맷으로 가공한다.
        /// 키: 테이블 명
        /// 값: 엑셀 시트에서 로드한 테이블 헤더 정보
        /// </summary>
        private readonly Dictionary<string, HeaderTable> headerTables = new Dictionary<string, HeaderTable>();

        /// <summary>
        /// 커스텀 열거형 테이블 컨테이너
        /// 키: 테이블 명
        /// 값: 코드 생성 룰을 무시한 특별 가공된 열거형 정보
        /// </summary>
        private readonly Dictionary<string, CustomEnumTable> customEnumTables = new Dictionary<string, CustomEnumTable>();

        /// <summary>
        /// 엑셀 파일을 로드하여 열거형과 헤더 테이블을 생성한다.
        /// </summary>
        /// <param name="excelPath">작업 파일 경로</param>
        /// <returns>성공 여부</returns>
        public bool Initialize(string excelPath)
        {
            if (true == string.IsNullOrEmpty(excelPath))
            {
                Console.WriteLine("Failed to generate. Excel path is null...");

                return false;
            }

            string[] files = Directory.GetFiles(excelPath);

            foreach (var file in files)
            {
                // 오픈 파일 제외
                if (true == Path.GetFileName(file).Contains("~$"))
                {
                    continue;
                }

                if (false == ExcelLoader.IsExcelExtension(Path.GetExtension(file)))
                {
                    continue;
                }

                string extendedProperties = string.Empty;

                var connection = ExcelLoader.GetOleDbConnectionStringBuilder(file);

                using (var context = new OleDbConnection(connection.ToString()))
                {
                    context.Open();

                    using (DataTable table = context.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" }))
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            string tableName = row["TABLE_NAME"].ToString().Trim('$');

                            if (tableName.Equals("enum_table", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            if (true == tableName.Contains('$'))
                            {
                                continue;
                            }

                            this.headerNames.Add(tableName);
                        }
                    }
                }
            }

            ExcelLoader.Initialize(excelPath, this);

            return true;
        }

        /// <summary>
        /// 해당 타입이 열거형 타입인지 확인한다.
        /// </summary>
        /// <param name="typeName">타입명</param>
        /// <returns>열거형 타입 여부</returns>
        public bool IsEnumerationTypes(string typeName)
        {
            foreach (var pair in this.enumTables)
            {
                if (true == pair.Value.IsExistsItem(typeName))
                {
                    return true;
                }
            }

            return false;
        }


        /// <summary>
        /// 해당 타입의 열거형 테이블을 얻는다.
        /// </summary>
        /// <param name="typeName">열거형 타입명</param>
        /// <returns>열거형 테이블</returns>
        public EnumTable GetEnumTableOrNull(string typeName)
        {
            foreach (var pair in this.enumTables)
            {
                if (true == pair.Value.IsExistsItem(typeName))
                {
                    return pair.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// 열거형 테이블을 관리 테이블에 추가한다.
        /// </summary>
        /// <param name="name">키</param>
        /// <param name="table">추가 테이블</param>
        /// <returns>성공 시 true</returns>
        public bool AddEnumTable(string name, EnumTable table)
        {
            if (true == this.enumTables.ContainsKey(name))
            {
                return false;
            }

            this.enumTables.Add(name, table);

            return this.enumTables.ContainsKey(name);
        }


        /// <summary>
        /// 열거형 테이블에 아이템을 추가한다.
        /// </summary>
        /// <param name="name">아이템 이름</param>
        /// <param name="item">in memory 데이터 테이블의 (row, col)에 위치하는 데이터</param>
        /// <param name="row">in memory 데이터 테이블의 행</param>
        /// <param name="col">in memory 데이터 테이블의 열</param>
        public void AddItemToEnumTable(string name, string item, int row, int col)
        {
            EnumTable temp;
            if (true == this.enumTables.TryGetValue(name, out temp))
            {
                temp.AddItem(item, row, col);
            }
        }



        /// <summary>
        /// 열거형 테이블에 멤버를 추가한다.
        /// </summary>
        /// <param name="name">아이템 이름</param>
        /// <param name="rowString">멤버 스트링</param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void AddMemberToEnumTable(string name, string rowString, int row, int col)
        {
            EnumTable temp;
            if (true == this.enumTables.TryGetValue(name, out temp))
            {
                temp.AddMember(rowString, row, col);
            }
        }


        /// <summary>
        /// 헤더 테이블을 파싱하기 위해 테이블 객체를 컨테이너에 생성한다.
        /// </summary>
        /// <param name="name">테이블명</param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="colMaxCount">읽어 들이는 테이블의 최대 컬럼 수, 엑셀 셀의 열 수</param>
        /// <param name="comments">멤버 주석 파싱 정보</param>
        public void CreateHeaderTable(string name, int row, int col, int colMaxCount, List<string> comments)
        {
            var headerTable = new HeaderTable
            {
                FileName = name,
                IDColumnRowIndex = row,
                IDColumnColIndex = col,
                EndColIndex = 0,
                
            };
            headerTable.AddRangeComments(comments);
            headerTable.AddColumn("ID", row, col, colMaxCount, this);

            this.headerTables.Add(name, headerTable);
        }


        public void AddColumnToHeaderTable(string name, string cellString, int row, int col, int maxColCount)
        {
            HeaderTable temp;
            if (true == this.headerTables.TryGetValue(name, out temp))
            {
                temp.AddColumn(cellString, row, col, maxColCount, this);
            }
        }


        public void AddCustomEnumTable(string name, CustomEnumTable table)
        {
            this.customEnumTables.Add(name, table);
        }


        /// <summary>
        /// 해당 경로에 코드를 생성한다.
        /// </summary>
        /// <param name="codePath">코드 생성 경로</param>
        public IEnumerable<string> GenerateCode(string codePath)
        {
            Console.WriteLine("\n");

            foreach (var pair in this.enumTables)
            {
                yield return pair.Value.GenerateCode(codePath);
            }

            foreach (var pair in this.customEnumTables)
            {
                yield return pair.Value.GenerateCode(codePath);
            }

            foreach (var pair in this.headerTables)
            {
                yield return pair.Value.GenerateCode(codePath);
            }

        }


        /// <summary>
        /// 해당 경로에 바이너리 파일을 생성한다.
        /// </summary>
        /// <param name="binaryPath">파일 생성 경로</param>
        public IEnumerable<string> GenerateBinary(string binaryPath)
        {
            Console.WriteLine("\n");

            foreach (var pair in this.headerTables)
            {
                if (true == string.Equals("LanguagePack", pair.Value.FileName.Trim('$', '_'), StringComparison.OrdinalIgnoreCase))
                {
                    yield return pair.Value.GenerateBinary(binaryPath);
                }
                else
                {
                    yield return pair.Value.GenerateBinary(binaryPath, "222B0F04E01040BB9B863EDDC7D2A431");
                }
            }
        }
    }
}
