/// 
/// Author - Mr Dk.
/// Version - 2019/08/28
/// Description - 
///     
///     The basic class of data structure for storing passwords
/// 

namespace Office.Utils.PasswdStructure
{
    abstract class PasswordStore
    {
        /// <summary>
        ///     Is there any more password?
        /// </summary>
        /// <returns></returns>
        public abstract bool HasNext();

        /// <summary>
        ///     Get the next password
        /// </summary>
        /// <returns></returns>
        public abstract string Next();

        /// <summary>
        ///     Reset the status of password store
        /// </summary>
        public abstract void Reset();
    }
}
