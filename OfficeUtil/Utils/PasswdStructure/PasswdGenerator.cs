///
/// Author - Mr Dk.
/// Version - 2019/08/29
/// Description -
///     Generate password of specific length from specific character set
///     
/// Usage - 
///     
///     PasswdGenerator gen = new PasswdGenerator("0123456789", 4);
///     while (gen.hasNext()) {
///         Console.WriteLine(gen.next());
///     }
/// 

using System.Text;

namespace Office.Utils.PasswdStructure
{
    class PasswdGenerator: PasswordStore
    {
        private string legalCharacters;
        private int length;
        private int cin; // 进位
        private int[] current; // index of character in legalCharacters

        public PasswdGenerator (string legalCharacterSet, int length)
        {
            this.legalCharacters = legalCharacterSet;
            this.length = length;
            this.current = new int[length];

            this.Reset();
        }

        public override void Reset()
        {
            this.cin = 0;
            for (int i = 0; i < length; i++)
            {
                this.current[i] = 0;
            }
        }

        public override bool HasNext()
        {
            return this.cin != 1 && this.legalCharacters.Length != 0 && this.length != 0;
        }

        public override string Next()
        {
            if (this.cin == 1 || this.length == 0 || this.legalCharacters.Length == 0)
            {
                return null;
            }

            StringBuilder strBuild = new StringBuilder();
            for (int i = 0; i < this.length; i++)
            {
                strBuild.Append(this.legalCharacters[this.current[i]]);
            }

            this.current[0] += 1;

            for (int i = 0; i < this.length; i++)
            {
                this.current[i] += this.cin;
                this.cin = 0;
                if (this.current[i] >= this.legalCharacters.Length)
                {
                    this.current[i] = 0;
                    this.cin = 1;
                }
                else
                {
                    break;
                }
            }

            return strBuild.ToString();
        }
    }
}
