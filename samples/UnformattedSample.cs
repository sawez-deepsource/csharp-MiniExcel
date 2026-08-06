using System;
using   System.Collections.Generic;
using System.Linq;

namespace   MiniExcelLibs.Samples
{
    public class    UnformattedSample
    {
        private   int _count=0;
            private string   _name = "sample" ;

        public UnformattedSample(   string name ,int count){
                _name=name;
            _count =count;
        }

        public int   Add(int a,int b )   {
        return a+b   ;
        }

        public   IEnumerable<int> EvenNumbers( List<int> numbers ){
            var  result=new List<int>() ;
                foreach(var n in numbers){
                    if(n%2==0)   {
                    result.Add( n ) ;
                    }
                }
            return result ;
        }

        public void   PrintAll(List<int>   items)
        {
                foreach ( var i in items )
            {
            Console.WriteLine(   i );
                }
        }

        public string Describe(){
          return $"{_name} has {_count} items"    ;
        }
    }
}
