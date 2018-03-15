using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MTools2013
{
    class GroupPlotter
    {
        private int[] _sectionArray;
        private int _noOfSections;
        private List<int> _lastIndexOfSection = new List<int>();
        public int NoOfSections 
        {
            get
            {
                for (int i = 0; i < _sectionArray.Length-1; i++ )
                {
                    if(_sectionArray[i] != _sectionArray[i+1] )
                    {
                        _lastIndexOfSection.Add(i);
                    }
                }

                _noOfSections = _lastIndexOfSection.Count;
                return _noOfSections;
            }
        }


        public GroupPlotter(int[] elements)
        {
            _sectionArray = elements;
        }
        public GroupPlotter()
        {

        }
    }
}
