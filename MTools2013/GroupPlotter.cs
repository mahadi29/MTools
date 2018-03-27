using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MTools2013
{
    class GroupPlotter
    {
        private string[] _sectionArray;
        private int _noOfSections;
        private List<int> _lastIndexOfSections = new List<int>();
        private List<int> _firstIndexOfSections = new List<int>();
        private List<string> _sectionNames = new List<string>();

        public List<int> FirstIndexOfSections
        {
            get
            {
                for (int i = 0; i < _sectionArray.Length - 1; i++)
                {
                    if (_sectionArray[i + 1] != "")
                    {
                        if (_sectionArray[i] == _sectionArray[i + 1] & i == 0)
                        {
                            _firstIndexOfSections.Add(i+1);
                        }
                        else if (_sectionArray[i] != _sectionArray[i + 1])
                        {
                            _firstIndexOfSections.Add(i + 2);
                        }
                    }
                }
                return _firstIndexOfSections;
            }
        }

        public List<int> LastIndexOfSections
        {
            get
            {
                for (int i = 0; i < _sectionArray.Length - 1; i++)
                {
                    if (_sectionArray[i] != "")
                    {
                        if (_sectionArray[i] != _sectionArray[i + 1])
                        {
                            _lastIndexOfSections.Add(i + 1);
                        }
                        else if (i + 1 == _sectionArray.Length - 1 & _sectionArray[i] == _sectionArray[i + 1])
                        {
                            _lastIndexOfSections.Add(i + 2);
                        }
                    }
                }
                return _lastIndexOfSections;
            }
        }
        public int NoOfSections 
        {
            get
            {
                if(_lastIndexOfSections.Count == _firstIndexOfSections.Count)
                {
                    _noOfSections = _lastIndexOfSections.Count;
                }
                return _noOfSections;
            }
        }
        public List<string> SectionNames
        {
            get
            {
                if (NoOfSections != 0)
                {
                    for (int i = 0; i < NoOfSections; i++ )
                    {
                        _sectionNames.Add(_sectionArray[_firstIndexOfSections.ElementAt(i)]);
                    }
                }
                else
                {
                    MessageBox.Show("");
                }
                return _sectionNames;
            }
        }

        public GroupPlotter(string[] SectionArray)
        {
            _sectionArray = SectionArray;
        }
        public GroupPlotter()
        {

        }
    }
}
