﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversionApp
{
    public class CellMap
    {
        public CellMap(int importId, int conversionId, string mapName)
        {
            ImportedCellId = importId;
            ConversionCellId = conversionId;
            MapName = mapName;
        }

        // the cell's Id in the file where the data is pulled from
        private int importedCellId;
        public int ImportedCellId
        {
            get { return importedCellId; }
            set
            {
                importedCellId = value;
            }
        }

        // the cell's ID for where the data should be newly stored
        private int conversionCellId;
        public int ConversionCellId
        {
            get { return conversionCellId; }
            set
            {
                conversionCellId = value;
            }
        }

        private string mapName;
        public string MapName
        {
            get { return mapName; }
            set
            {
                mapName = value;
            }
        }

    }
}
