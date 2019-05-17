using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meridium.Excel
{
    public class ExcelProperties
    {

        private bool useHeaders;
        private bool useMex;
        private bool isText;

        public bool UseHeaders
        {
            get { return useHeaders; }
            set { useHeaders = value; }
        }


        public bool UseMex
        {
            get { return useMex; }
            set { useMex = value; }
        }


        public bool IsText
        {
            get { return isText; }
            set { isText = value; }
        }

        /// <summary>
        /// Defaults: UseHeaders=true, UseMex=false, IsText=false;
        /// </summary>
        /// <remarks>
        /// Warning! UseMex (IMEX) must be false if you are creating a file (or table).
        /// </remarks>
        public ExcelProperties()
        {
            useHeaders = true;
            useMex = false;
            isText = false;
        }

        /// <summary>
        /// No defaults.
        /// </summary>
        /// <remarks>
        /// Warning! UseMex (IMEX) must be false if you are creating a file (or table).
        /// </remarks>
        public ExcelProperties(bool useHeaders, bool useMex, bool isText)
        {
            this.useHeaders = useHeaders;
            this.useMex = useMex;
            this.isText = isText;
        }


    }
}
