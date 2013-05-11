using System;
using System.Collections.Generic;
using System.Web;

namespace FlexService.model
{
    public class AutoDLN
    {
        private string signA;

        public string SignA
        {
            get { return signA; }
            set { signA = value; }
        }

        private string signB;

        public string SignB
        {
            get { return signB; }
            set { signB = value; }
        }

        private string lineIndex;

        public string LineIndex
        {
            get { return lineIndex; }
            set { lineIndex = value; }
        }

        private string cardCode;

        public string CardCode
        {
            get { return cardCode; }
            set { cardCode = value; }
        }

        private string itemCode;

        public string ItemCode
        {
            get { return itemCode; }
            set { itemCode = value; }
        }

        private string itemName;

        public string ItemName
        {
            get { return itemName; }
            set { itemName = value; }
        }

        private string fa1;

        public string Fa1
        {
            get { return fa1; }
            set { fa1 = value; }
        }

        private string fa2;

        public string Fa2
        {
            get { return fa2; }
            set { fa2 = value; }
        }

        private string fa3;

        public string Fa3
        {
            get { return fa3; }
            set { fa3 = value; }
        }

        private string fa4;

        public string Fa4
        {
            get { return fa4; }
            set { fa4 = value; }
        }

        private string sum;

        public string Sum
        {
            get { return sum; }
            set { sum = value; }
        }

        private string wH;

        public string WH
        {
            get { return wH; }
            set { wH = value; }
        }

        private List<AutoDLN1> lines = new List<AutoDLN1>();

        public List<AutoDLN1> Lines
        {
            get { return lines; }
            set { lines = value; }
        }
    }
}