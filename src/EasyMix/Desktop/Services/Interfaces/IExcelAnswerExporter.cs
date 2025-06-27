﻿using Desktop.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Desktop.Services.Interfaces
{
    public interface IExcelAnswerExporter
    {
        public void ExportExcelAnswers(string filePath, List<QuestionExport> answers);
    }
}
