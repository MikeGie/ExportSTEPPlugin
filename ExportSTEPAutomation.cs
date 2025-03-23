/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

using Inventor;
using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using Shared;

namespace ExportSTEPPlugin
{
    [ComVisible(true)]
    public class ExportSTEPAutomation : AutomationBase
    {
        /// <summary>
        /// File name of output STEP file.
        /// The file name is expected by the corresponding Activity.
        /// </summary>
        private const string OutputStepName = "output.step";

        public ExportSTEPAutomation(InventorServer inventorApp) : base(inventorApp)
        {
        }

        public override void Run(Document doc)
        {
            LogTrace("Processing " + doc.FullFileName);

            try
            {
                switch (doc.DocumentType)
                {
                    case DocumentTypeEnum.kPartDocumentObject:
                        ProcessPart((PartDocument)doc);
                        break;

                    case DocumentTypeEnum.kAssemblyDocumentObject:
                        ProcessAssembly((AssemblyDocument)doc);
                        break;

                    // complain about non-supported document types
                    default:
                        throw new ArgumentOutOfRangeException(nameof(doc), "Unsupported document type");
                }

                LogTrace("STEP export completed successfully.");
            }
            catch (Exception e)
            {
                LogError("Processing failed. " + e.ToString());
            }
        }

        public override void ExecWithArguments(Document doc, NameValueMap map)
        {
            LogError("Unexpected execution path! ExportSTEP does not expect any extra arguments!");
        }

        private void ProcessPart(PartDocument doc)
        {
            using (new HeartBeat())
            {
                LogTrace("Exporting Part document to STEP format");

                // Create STEP export options
                TranslatorAddIn stepTranslator = GetStepTranslator();
                if (stepTranslator == null)
                {
                    LogError("Could not get STEP translator.");
                    return;
                }

                TranslationContext context = _inventorApplication.TransientObjects.CreateTranslationContext();
                context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                NameValueMap options = _inventorApplication.TransientObjects.CreateNameValueMap();
                options.Value["ApplicationProtocolType"] = 3; // AP214 - automotive design
                options.Value["Author"] = System.Environment.UserName;
                options.Value["Organization"] = "Autodesk";

                DataMedium dataMedium = _inventorApplication.TransientObjects.CreateDataMedium();
                dataMedium.FileName = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), OutputStepName);

                // Perform the export
                stepTranslator.SaveCopyAs(doc, context, options, dataMedium);

                LogTrace($"STEP file exported successfully to {dataMedium.FileName}");
            }
        }

        private void ProcessAssembly(AssemblyDocument doc)
        {
            using (new HeartBeat())
            {
                LogTrace("Exporting Assembly document to STEP format");

                // Create STEP export options
                TranslatorAddIn stepTranslator = GetStepTranslator();
                if (stepTranslator == null)
                {
                    LogError("Could not get STEP translator.");
                    return;
                }

                TranslationContext context = _inventorApplication.TransientObjects.CreateTranslationContext();
                context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                NameValueMap options = _inventorApplication.TransientObjects.CreateNameValueMap();
                options.Value["ApplicationProtocolType"] = 3; // AP214 - automotive design
                options.Value["Author"] = System.Environment.UserName;
                options.Value["Organization"] = "Autodesk";

                DataMedium dataMedium = _inventorApplication.TransientObjects.CreateDataMedium();
                dataMedium.FileName = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), OutputStepName);

                // Perform the export
                stepTranslator.SaveCopyAs(doc, context, options, dataMedium);

                LogTrace($"STEP file exported successfully to {dataMedium.FileName}");
            }
        }

        private TranslatorAddIn GetStepTranslator()
        {
            string stepTranslatorId = "STEP_Translator";

            ApplicationAddIns appAddIns = _inventorApplication.ApplicationAddIns;
            TranslatorAddIn stepTranslator = null;

            foreach (ApplicationAddIn addIn in appAddIns)
            {
                if (addIn.ClassIdString == stepTranslatorId && addIn.Activated)
                {
                    stepTranslator = (TranslatorAddIn)addIn;
                    break;
                }
            }

            if (stepTranslator == null)
            {
                LogTrace("STEP translator not found or not activated.");
            }

            return stepTranslator;
        }
    }
}