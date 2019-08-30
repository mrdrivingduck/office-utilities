///
/// Author - Mr Dk.
/// Version - 2019/08/30
/// Description -
///     Try all passwords in a specific password array
/// 

using System.Collections.Generic;

namespace Office.Utils.PasswdStructure
{
    class PasswordUnion : PasswordStore
    {
        private List<string> array;
        private int current;

        public PasswordUnion(List<string> passwords)
        {
            this.array = new List<string>(passwords);
            this.current = 0;
        }

        /// <summary>
        ///     Move the current password to the head
        ///     Reset the current pointer
        /// </summary>
        public override void Reset()
        {
            if (this.array.ToArray().Length == 0)
            {
                return;
            }
            if (this.current <= this.array.ToArray().Length &&
                this.current - 1 > 0)
            {
                string selected = this.array[this.current-1];
                this.array.Remove(this.array[this.current-1]);
                this.array.Insert(0, selected);
            }
            this.current = 0;
        }

        /// <summary>
        ///     Whether the current pointer reaches the end
        /// </summary>
        /// <returns></returns>
        public override bool HasNext()
        {
            int length = this.array.ToArray().Length;
            if (length == 0)
            {
                return false;
            }
            if (current < length)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        ///     Get the current password and turn to the next
        /// </summary>
        /// <returns></returns>
        public override string Next()
        {
            string passwd = this.array[this.current];
            this.current++;
            return passwd;
        }
    }
}
