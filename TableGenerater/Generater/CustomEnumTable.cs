using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Generater
{
    /// <summary>
    /// 엑셀 테이블의 ID가 멤버인 열거형 테이블
    /// </summary>
    public class CustomEnumTable
    {
        /// <summary>
        /// 파일명
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 테이블에서 주석 위치
        /// </summary>
        public int CommentIndex { get; set; }

        /// <summary>
        /// 멤버명
        /// </summary>
        public readonly List<string> memberNames = new List<string>();

        /// <summary>
        /// 멤버 정수형 데이터
        /// </summary>
        public readonly List<int> memberValues = new List<int>();

        /// <summary>
        /// 주석
        /// </summary>
        public readonly List<string> comments = new List<string>();


        public void SetData(List<Tuple<string, string>> info)
        {
            int index = 0;
            foreach (var pair in info)
            {
                this.memberNames.Add(pair.Item1);
                this.memberValues.Add(index);
                this.comments.Add(pair.Item2);

                ++index;
            }

            this.memberNames.Add("Max");
            this.memberValues.Add(index);
            this.comments.Add(string.Empty);
        }

        public string GenerateCode(string headerPath)
        {
            string enumName = "TableEnum_" + this.FileName.Trim('$');
            var codeEnumNamespace = new System.CodeDom.CodeNamespace("TableHeader");

            if (this.memberNames.Count != this.memberValues.Count)
            {
                Console.WriteLine("{0} memberName Count:{1}, memberValue Count:{2}", this.FileName, this.memberNames.Count, this.memberValues.Count);
            }

            var enumClass = new System.CodeDom.CodeTypeDeclaration(enumName)
            {
                Attributes = System.CodeDom.MemberAttributes.Public,
                IsEnum = true
            };

            for (var i = 0; i < this.memberNames.Count; ++i)
            {
                var memberField = new System.CodeDom.CodeMemberField(typeof(int), this.memberNames[i]);
                memberField.InitExpression = new System.CodeDom.CodePrimitiveExpression(this.memberValues[i]);

                // 가장 우측에 정의된 열거형의 모든 멤버에 주석이 없는 경우 해당 컬럼은 클립핑 영역에서 제외되어 
                // 주석을 공란으로 읽어올 수 없다.
                if (0 < this.comments.Count &&
                    false == string.IsNullOrEmpty(this.comments[i]))
                {
                    memberField.Comments.Add(new System.CodeDom.CodeCommentStatement("<summary>", true));
                    memberField.Comments.Add(new System.CodeDom.CodeCommentStatement(this.comments[i], true));
                    memberField.Comments.Add(new System.CodeDom.CodeCommentStatement("</summary>", true));
                }

                enumClass.Members.Add(memberField);
            }

            codeEnumNamespace.Types.Add(enumClass);

            var filePath = string.Format(@"{0}\{1}.cs", headerPath, enumName);
            File.Delete(filePath);

            CodeGenerater.CrateCodeFile(filePath, codeEnumNamespace);

            Console.WriteLine("Create code: {0}", filePath);

            return "Create code: " + filePath;
        }
    }
}
