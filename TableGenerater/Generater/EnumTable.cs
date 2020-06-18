using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Generater
{
    /// <summary>
    /// 엑셀에서 로드한 열거형 정보로 코드 생성시 해당 객체의 숫자만큼 파일이 생성된다.
    /// </summary>
    /// 
    /// <warning>
    /// 테이블 헤더에 타입으로 사용하는 열거형은 테이블 정의 열거형만 사용할 수 있다.
    /// </warning>
    public class EnumTable
    {
        /// <summary>
        /// 플래그 속성 토큰
        /// </summary>
        public const string TOKEN_FLAGS = "Flags_";


        /// <summary>
        /// 하나의 열거형 객체
        /// </summary>
        private class Item
        {
            /// <summary>
            /// 플래그 속성 여부
            /// </summary>
            public bool isFlags;

            /// <summary>
            /// 열거형 이름
            /// </summary>
            public string name;
            
            /// <summary>
            /// in memory 데이터 테이블에서의 상대적 위치 행
            /// </summary>
            public int startRowIndex;

            /// <summary>
            /// in memory 데이터 테이블에서의 상대적 위치 열
            /// </summary>
            public int endRowIndex;
            /// <summary>
            /// 열거형 멤버명이 정의된 인덱스 (zero-base 값)
            /// </summary>
            public int nameColumnIndex;
            /// <summary>
            /// 열거형 멤버의 정수형 값이 정의된 인덱스
            /// 엑셀에서 열거형 멤버명의 우측 컬럼에 정의한다.
            /// </summary>
            public int valueColumnIndex;

            /// <summary>
            /// 주석이 정의된 인덱스
            /// 엑셀에서 열거형 숫자 정의 우측 컬럼에 정의한다.
            /// </summary>
            public int commentsIndex;

            public readonly List<string> memberNames = new List<string>();
            public readonly List<long> memberValues = new List<long>();
            public readonly List<string> comments = new List<string>();
        }

        /// <summary>
        /// 파일명
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 파일에 정의되는 열거형 정보
        /// </summary>
        private readonly Dictionary<string, Item> items = new Dictionary<string, Item>();

        /// <summary>
        /// 열거형 아이템 포함 여부 확인
        /// </summary>
        /// <param name="typeName">열거형 타입명</param>
        /// <returns>포함 여부</returns>
        public bool IsExistsItem(string typeName)
        {
            return this.items.ContainsKey(typeName);
        }


        /// <summary>
        /// 해당 열거형 타입의 멤버 여부를 확인한다.
        /// </summary>
        /// <param name="typeName">열거형 타입명</param>
        /// <param name="memberName">확인할 멤버명</param>
        /// <returns>멤버 여부</returns>
        public bool IsMember(string typeName, string memberName)
        {
            Item temp;
            if (true == this.items.TryGetValue(typeName, out temp))
            {
                if (true == temp.isFlags)
                {
                    var split = memberName.Split('|');

                    for (int i = 0; i < split.Length; ++i)
                    {
                        if (false == temp.memberNames.Contains(split[i].Trim()))
                        {
                            return false;
                        }
                    }

                    return true;
                }
                else
                {
                    return temp.memberNames.Contains(memberName);
                }
            }

            return false;
        }

        /// <summary>
        /// 아이템을 추가한다.
        /// 아이템은 하나의 열거형 객체가 된다.
        /// </summary>
        /// <param name="name">열거형명</param>
        /// <param name="row">in memory 데이터 테이블 행</param>
        /// <param name="col">in memory 데이터 테이블 열</param>
        public void AddItem(string name, int row, int col)
        {
            this.items.Add(name, new Item
                {
                    isFlags = name.Contains(TOKEN_FLAGS),
                    name = name,
                    startRowIndex = row,
                    nameColumnIndex = col,
                    valueColumnIndex = col + 1,
                    commentsIndex = col + 2
            });
        }

        
        /// <summary>
        /// 열거형 멤버명과 값을 추가한다.
        /// 해당 함수를 호출하기 전에 아이템이 생성된 상태여야한다.
        /// </summary>
        /// <param name="rowString">파싱 값</param>
        /// <param name="row">in memory 데이터 테이블 행</param>
        /// <param name="col">in memory 데이터 테이블 열</param>
        public void AddMember(string rowString, int row, int col)
        {
            try
            {
                foreach (var pair in items)
                {
                    if (row <= pair.Value.startRowIndex 
                        || (0 != pair.Value.endRowIndex && pair.Value.endRowIndex <= row))
                    {
                        continue;
                    }

                    if (col == pair.Value.nameColumnIndex)
                    {
                        if (!rowString.Any())
                        {
                            if (0 == pair.Value.endRowIndex)
                            {
                                pair.Value.endRowIndex = row;
                            }

                            continue;
                        }

                        if (pair.Value.memberNames.Contains(rowString))
                        {
                            throw new InvalidOperationException(string.Format("EnumClass MemberName Duplicated TableTabName:{0} EnumClassName:{1} MemberName:{2}", this.FileName, pair.Value.name, rowString));
                        }

                        pair.Value.memberNames.Add(rowString);
                        break;
                    }
                    else if (col == pair.Value.valueColumnIndex)
                    {
                        if (!rowString.Any())
                        {
                            if (0 == pair.Value.endRowIndex)
                            {
                                pair.Value.endRowIndex = row;
                            }

                            continue;
                        }

                        pair.Value.memberValues.Add(Convert.ToInt64(rowString));
                        break;
                    }
                    else if (col == pair.Value.commentsIndex)
                    {
                        //if (!rowString.Any())
                        //{
                        //    //if (0 == pair.Value.endRowIndex)
                        //    //{
                        //    //    pair.Value.endRowIndex = row;
                        //    //}

                        //    continue;
                        //}

                        pair.Value.comments.Add(rowString);
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }


        /// <summary>
        /// 열거형 테이블을 이용하여 해당 경로에 소스 코드를 생성한다.
        /// </summary>
        /// <param name="headerPath">작업 경로</param>
        public string GenerateCode(string headerPath)
        {
            string enumName = "TableEnum_" + this.FileName.Trim('$');
            var codeEnumNamespace = new System.CodeDom.CodeNamespace("TableHeader");
            
            foreach (var pair in this.items)
            {
                if (pair.Value.memberNames.Count != pair.Value.memberValues.Count)
                {
                    Console.WriteLine("{0} memberName Count:{1}, memberValue Count:{2}", pair.Key, pair.Value.memberNames.Count, pair.Value.memberValues.Count);
                }

                var enumClass = new System.CodeDom.CodeTypeDeclaration(pair.Value.name)
                {
                    Attributes = System.CodeDom.MemberAttributes.Public,
                    IsEnum = true,
                };

                Type memberType;

                if (true == pair.Value.isFlags)
                {
                    enumClass.BaseTypes.Add("System.Int64");
                    enumClass.CustomAttributes.Add(new System.CodeDom.CodeAttributeDeclaration("System.FlagsAttribute"));

                    memberType = typeof(long);
                }
                else
                {
                    memberType = typeof(int);
                }

                for (var i = 0; i < pair.Value.memberNames.Count; ++i)
                {
                    var memberField = new System.CodeDom.CodeMemberField(memberType, pair.Value.memberNames[i]);
                    memberField.InitExpression = new System.CodeDom.CodePrimitiveExpression(pair.Value.memberValues[i]);

                    // 가장 우측에 정의된 열거형의 모든 멤버에 주석이 없는 경우 해당 컬럼은 클립핑 영역에서 제외되어 
                    // 주석을 공란으로 읽어올 수 없다.
                    if (0 < pair.Value.comments.Count &&
                        false == string.IsNullOrEmpty(pair.Value.comments[i]))
                    {
                        memberField.Comments.Add(new System.CodeDom.CodeCommentStatement("<summary>", true));
                        memberField.Comments.Add(new System.CodeDom.CodeCommentStatement(pair.Value.comments[i], true));
                        memberField.Comments.Add(new System.CodeDom.CodeCommentStatement("</summary>", true));
                    }

                    enumClass.Members.Add(memberField);
                }

                codeEnumNamespace.Types.Add(enumClass);
            }

            var filePath = string.Format(@"{0}\{1}.cs", headerPath, enumName);
            File.Delete(filePath);

            CodeGenerater.CrateCodeFile(filePath, codeEnumNamespace);

            Console.WriteLine("Create code: {0}", filePath);

            return "Create code: " + filePath;
        }
    }
}
