using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Mapping
{
    class tester
    {
        static void Fun(string[] args)
        {
            string[] points = { "taskas1", "taskas2" };
            string line=@" Macro 353 \taskas1\ \taskas2\";
            string tempStr, tempStr1;

            Console.WriteLine(line);
            if (line.Contains('\\'))//if point identifier is found
            {
                tempStr = line.Substring(line.IndexOf('\\')+1, line.Length - line.IndexOf('\\')-1);
                while (tempStr.Length > 0)
                {
                    if (points.Contains(tempStr.Substring(0, tempStr.IndexOf('\\'))))
                    {
                        Console.WriteLine("points: " + tempStr.Substring(0, tempStr.IndexOf('\\')));
                        
                        //check for line begining
                        if (!(line.IndexOf(' ') == 0))//if the line is primary
                        {
                            if (line.Substring(0, line.IndexOf(' ')).ToUpper()=="MACRO")
                            {
                                tempStr1=line.Substring(line.IndexOf(' ')+1,line.Length-line.IndexOf(' ')-1);
                                tempStr1 =line.Substring(line.IndexOf(' ') + 1, tempStr1.IndexOf(' '));
                                Console.WriteLine("MARCO " + tempStr1);
                            }
                        }                      
                    }
                     
                    //move to the next point identifier
                    if (tempStr.Contains('\\'))//if point identifier is found
                    {
                        
                        tempStr = tempStr.Substring(tempStr.IndexOf('\\') + 1, tempStr.Length - tempStr.IndexOf('\\') - 1);
                    }
                    else
                        break;

                    if (tempStr.Contains('\\'))//if point identifier is found
                        tempStr = tempStr.Substring(tempStr.IndexOf('\\')+1, tempStr.Length - tempStr.IndexOf('\\')-1);
                    else
                        break;
                    Console.WriteLine(tempStr);
                }

            }
        }
    }
}
