using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

using TableHeader;

namespace Generater
{
    /// <summary>
    /// 엑셀에서 로드한 테이블 헤더 정보로 코드 생성 시 단일 테이블 헤더를 생성한다.
    /// 멤버 클래스나 객체는 엑셀을 기준으로 이해할 필요가 있다.
    /// </summary>
    /// 
    /// <warning>
    /// 데이터 타입보다 컬럼명이 더 많은 경우 컬럼명을 제거한다. 즉 열의 기준은 데이터 타입이된다.
    /// 엑셀 데이터를 작성할 때 컬럼명과 데이터타입의 마지막 열 바로 다음 열은 비어있는 열로 만드는게 좋다. 비고 등은 한열을 비우고 작성하자.
    /// 테이블 헤더에 타입으로 사용하는 열거형은 테이블 정의 열거형만 사용할 수 있다.
    /// </warning>
    public class HeaderTable
    {
        /// <summary>
        /// 바이너리 값 파싱 토큰
        /// </summary>
        public const string TOKEN_BINARY_VALUE = ";";
        public const string TOKEN_ROW_BINARY_VALUE = "#";

        /// <summary>
        /// 컬럼 정보
        /// </summary>
        class ColumnInfo
        {
            public int ColumnIndex { get; set; }
            /// <summary>
            /// 테이블 정의형 열거형인 경우 타입이 null
            /// </summary>
            public Type DataType { get; set; }
            public string DataString { get; set; }
        }

        /// <summary>
        /// 테이블 헤더 멤버 타입
        /// </summary>
        private readonly List<ColumnInfo> memberTypes = new List<ColumnInfo>();

        /// <summary>
        /// 테이블 헤더 멤버명
        /// </summary>
        private readonly List<ColumnInfo> memberNames = new List<ColumnInfo>();

        /// <summary>
        /// 엑셀 셀 값
        /// </summary>
        private readonly List<string> cellValues = new List<string>();

        /// <summary>
        /// 스킵 스트링
        /// </summary>
        private readonly List<string> skipStrings = new List<string>();

        /// <summary>
        /// 멤버 주석
        /// </summary>
        private readonly List<string> memberComments = new List<string>();

        /// <summary>
        /// 라인 스킵 여부
        /// </summary>
        private bool IsSkipRow { get; set; }

        /// <summary>
        /// 스킵할 행 인덱스
        /// </summary>
        private int SkipRowIndex { get; set; }

        /// <summary>
        /// 테이블 헤더 파일명
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// in memory 데이터 테이블에서 컬럼명 ID가 존재하는 행 인덱스
        /// </summary>
        public int IDColumnRowIndex { get; set; }

        /// <summary>
        /// in memory 데이터 테이블에서 컬럼명 ID가 존재하는 열 인덱스
        /// </summary>
        public int IDColumnColIndex { get; set; }

        public int EndColIndex { get; set; }

        /// <summary>
        /// 멤버 주석을 추가한다.
        /// </summary>
        /// <param name="comments"></param>
        public void AddRangeComments(List<string> comments)
        {
            this.memberComments.AddRange(comments);
        }


        public void AddColumn(string data, int row, int col, int maxColCount, Headergenerater container)
        {
            #region 스킵 스트링 생성 및 스킵 확인
            if (false == this.skipStrings.Any())
            {
                for (int i = 0; i < maxColCount; ++i)
                {
                    this.skipStrings.Add(string.Format("F{0}", i));
                }
            }
            else
            {
                if (true == skipStrings.Contains(data))
                {
                    this.EndColIndex = col - 1;
                    return;
                }
            }
            #endregion

            if (0 != this.EndColIndex && this.EndColIndex < col)
            {
                return;
            }


            // 키워드 ID가 셋팅된 라인은 컬럼명이 정의되어 있고 컬럼명은 테이블 헤더 멤버명으로 사용된다.
            if (row == this.IDColumnRowIndex && this.IDColumnColIndex <= col)
            {
                if (false == data.Any())
                {
                    // 엑셀의 우측 방향으로 파싱해가다 비어있는 셀을 만나면 이전 셀을 마지막 컬럼으로 셋팅한다.
                    this.EndColIndex = col - 1;
                    return;
                }

                this.memberNames.Add(new ColumnInfo { DataString = data, ColumnIndex = col });
            }
            // 컬럼명 바로 아래 라인은 데이터 타입이 정의되어 있다.
            else if (row == this.IDColumnRowIndex + 1 && this.IDColumnColIndex <= col)
            {
                if (false == data.Any())
                {
                    this.EndColIndex = col - 1;

                    // 데이터 타입보다 컬럼명이 더 많은 경우 컬럼명을 제거한다.
                    // 즉 열의 기준은 데이터 타입이된다.
                    if (this.EndColIndex < this.memberNames.Count - 1)
                    {
                        var removeIndex = this.memberNames.Count;
                        var removeCount = this.memberNames.Count - (this.EndColIndex + 1);
                        while (removeCount != removeIndex)
                        {
                            --removeIndex;
                            this.memberNames.RemoveAt(removeIndex);
                        }
                    }
                    return;
                }

                Type type = CodeGenerater.ConvertToTypeOrNull(data);
                if (null == type && false == container.IsEnumerationTypes(data))
                {
                    throw new InvalidOperationException(string.Format("Not Allowed TableName : {0} Type : {1} Column : {2} Row : {3}", this.FileName, data, col + 1, row + 1));
                }

                this.memberTypes.Add(new ColumnInfo { ColumnIndex = col, DataString = data, DataType = type });
            }
            // 데이터 타입명 아래 라인부터는 데이터 셀이다.
            else if (this.IDColumnRowIndex + 1 < row && this.IDColumnColIndex <= col)
            {
                // 공백 행 확인
                if (this.IDColumnColIndex == col)
                {
                    if (false == data.Any())
                    {
                        this.IsSkipRow = true;
                        this.SkipRowIndex = row;
                    }
                    else
                    {
                        this.IsSkipRow = false;
                    }
                }

                // 공백 행 파싱 스킵
                if (true == this.IsSkipRow && row == this.SkipRowIndex)
                {
                    return;
                }

                ColumnInfo columnInfo = this.memberTypes.Find(p => p.ColumnIndex == col);
                if (null == columnInfo.DataType)
                {
                    EnumTable enumTable = container.GetEnumTableOrNull(columnInfo.DataString);
                    if (null == enumTable)
                    {
                        throw new InvalidOperationException(string.Format("Not Allowed TableName : {0}, data : {1}, Column : {2}, Row : {3}", this.FileName, data, col + 1, row + 1));
                    }

                    if (false == enumTable.IsMember(columnInfo.DataString, data))
                    {
                        throw new InvalidOperationException(string.Format("Not Allowed TableName : {0}, data : {1}, Column : {2}, Row : {3}", this.FileName, data, col + 1, row + 1));
                    }
                }

                this.cellValues.Add(data);
            }
        }


        /// <summary>
        /// 헤더 테이블을 이용하여 해당 경로에 헤더 테이블 클래스 파일을 만든다.
        /// </summary>
        /// <param name="headerPath">작업 경로</param>
        public string GenerateCode(string headerPath)
        {
            string headerName = "TableHeader_" + this.FileName.Trim('$');
            var codeNamespace = new System.CodeDom.CodeNamespace("TableHeader");

            var headerClass = new System.CodeDom.CodeTypeDeclaration(headerName)
            {
                Attributes = System.CodeDom.MemberAttributes.Public
            };

            for (var i = 0; i < this.memberNames.Count; ++i)
            {
                var attributeName = this.memberNames[i].DataString + " { get; set; }//";

                System.CodeDom.CodeMemberField memberField = null;

                if (null != this.memberTypes[i].DataType)
                {
                    memberField = new System.CodeDom.CodeMemberField(this.memberTypes[i].DataType, attributeName);
                }
                else
                {
                    memberField = new System.CodeDom.CodeMemberField(this.memberTypes[i].DataString, attributeName);
                }

                // 주석이 생략된 경우도 있다.
                if (0 < this.memberComments.Count && false == string.IsNullOrEmpty(this.memberComments[i]))
                {
                    memberField.Comments.Add(new System.CodeDom.CodeCommentStatement("<summary>", true));
                    memberField.Comments.Add(new System.CodeDom.CodeCommentStatement(this.memberComments[i], true));
                    memberField.Comments.Add(new System.CodeDom.CodeCommentStatement("</summary>", true));
                }

                memberField.Attributes = System.CodeDom.MemberAttributes.Public | System.CodeDom.MemberAttributes.Final;
                headerClass.Members.Add(memberField);
            }

            codeNamespace.Types.Add(headerClass);

            var filePath = string.Format(@"{0}\{1}.cs", headerPath, headerName);

            File.Delete(filePath);

            CodeGenerater.CrateCodeFile(filePath, codeNamespace);

            Console.WriteLine("Create code: {0}", filePath);

            return "Create Code: " + filePath;
        }


        /// <summary>
        /// 헤더 테이블을 이용하여 데이터 바이너리를 만든다.
        /// </summary>
        /// <param name="binaryPath">작업 경로</param>
        public string GenerateBinary(string binaryPath)
        {
            var fileName = string.Format(@"{0}\{1}.bytes", binaryPath, this.FileName.Trim('$', '_'));

            File.Delete(fileName);

            int columnCount = memberTypes.Count;

            using (var sw = new StreamWriter(new FileStream(fileName, FileMode.Create)))
            {
                for (var i = 0; i < this.cellValues.Count; ++i)
                {
                    sw.Write(this.cellValues[i]);

                    // 세미콜론은 토큰이다.
                    if (i != this.cellValues.Count - 1)
                    {
                        sw.Write( ( (i+1) % columnCount == 0) ? TOKEN_ROW_BINARY_VALUE : TOKEN_BINARY_VALUE);
                    }
                }
            }

            Console.WriteLine("Create binary: {0}", fileName);
            return "Create bynary: " + fileName;
        }

        public string GenerateBinary(string binaryPath, string encryptionKey)
        {
            var fileName = string.Format(@"{0}\{1}.bytes", binaryPath, this.FileName.Trim('$', '_'));

            File.Delete(fileName);

            int columnCount = memberTypes.Count;


            using (var ms = new MemoryStream())
            {
                using (var sw = new StreamWriter(ms))
                {
                    for (var i = 0; i < this.cellValues.Count; ++i)
                    {
                        sw.Write(this.cellValues[i]);

                        // 세미콜론은 토큰이다.
                        if (i != this.cellValues.Count - 1)
                        {
                            sw.Write(((i + 1) % columnCount == 0) ? TOKEN_ROW_BINARY_VALUE : TOKEN_BINARY_VALUE);
                        }
                    }
                }

                // 암호화한다.
                byte[] readBytes = Encryption.AESEncrypt128(ms.ToArray(), encryptionKey);

                var binaryFormatter = new BinaryFormatter();
                var file = new FileInfo(fileName);

                using (var binaryFile = file.Create())
                {
                    binaryFormatter.Serialize(binaryFile, readBytes);
                    binaryFile.Flush();
                }
            }

            Console.WriteLine("Create binary: {0}", fileName);
            return "Create bynary: " + fileName;
        }
    }
}
