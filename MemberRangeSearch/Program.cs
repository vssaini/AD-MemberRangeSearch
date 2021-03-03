using System;
using System.Collections.Generic;
using System.DirectoryServices;

namespace MemberRangeSearch
{
    class Program
    {
        static void Main()
        {
            var users = RetrieveMembers("LDAP://CN=1000PlusTest,OU=Groups,DC=domain,DC=com");

            // Show user's details
            foreach (var user in users)
            {
                Console.WriteLine(user);
            }

            Console.WriteLine("");
            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("Total {0} users retrieved.", users.Count);
            Console.WriteLine("-----------------------------------------------");

            Console.ReadKey();
        }

        /// <summary>
        /// Retrieve details for attribute 'member' for specific group.
        /// </summary>
        /// <param name="ldapPath">The LDAP path for group.</param>
        /// <returns>Return value for member else null.</returns>
        private static HashSet<string> RetrieveMembers(string ldapPath)
        {
            var users = new HashSet<string>();

            using (var dirEntry = new DirectoryEntry(ldapPath))
            {
                using (var dirSearcher = new DirectorySearcher(dirEntry) { Filter = "(objectClass=*)" })
                {
                    const uint rangeStep = 1000;
                    uint rangeLow = 0;
                    uint rangeHigh = rangeLow + (rangeStep - 1);
                    bool lastQuery = false;
                    bool quitLoop = false;

                    do
                    {
                        // 1. Set attribute with range
                        var attributeWithRange = !lastQuery ? String.Format("member;range={0}-{1}", rangeLow, rangeHigh) : String.Format("member;range={0}-*", rangeLow);

                        // 2. Clear properties to load and add afresh
                        dirSearcher.PropertiesToLoad.Clear();
                        dirSearcher.PropertiesToLoad.Add(attributeWithRange);

                        // 3. Retrieve result
                        SearchResult results = dirSearcher.FindOne();

                        // 4. If result contains attribute, hold details
                        if (results.Properties.Contains(attributeWithRange))
                        {
                            foreach (var obj in results.Properties[attributeWithRange])
                            {
                                var user = Convert.ToString(obj);
                                if (!users.Contains(user))
                                    users.Add(user);
                            }

                            // If last query was set, then time to quit loop
                            if (lastQuery) quitLoop = true;
                        }
                        else
                        {
                            // For handling empty group to avoid infinite loop
                            if (lastQuery == false)
                            {
                                lastQuery = true;
                            }
                            else
                            {
                                quitLoop = true;
                            }
                        }

                        // 5. If it was not last query, then we need to further modify ranges
                        if (!lastQuery)
                        {
                            rangeLow = rangeHigh + 1;
                            rangeHigh = rangeLow + (rangeStep - 1);
                        }
                    }
                    while (!quitLoop);
                }
            }

            return users;
        }
    }
}
