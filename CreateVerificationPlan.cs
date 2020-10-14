////////////////////////////////////////////////////////////////////////////////
// CreateVerificationPlan.cs
//
//  A ESAPI v13.6+ script that demonstrates creation of verification plans 
//  from a clinical plan.
//
// Kata Intermediate.5)    
//  Program an ESAPI automation script that creates a new QA course, a new set 
//  of verification plans for the selected clinical plan 
//  (1 composite and 1 verification plan per beam), and calculates dose for all 
//  of the new verification plans.
//
// Applies to:
//      Eclipse Scripting API for Research Users
//          13.6, 13.7, 15.0,15.1
//      Eclipse Scripting API
//          15.1
//
// Copyright (c) 2017 Varian Medical Systems, Inc.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy 
// of this software and associated documentation files (the "Software"), to deal 
// in the Software without restriction, including without limitation the rights 
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
// copies of the Software, and to permit persons to whom the Software is 
// furnished to do so, subject to the following conditions:
//
//  The above copyright notice and this permission notice shall be included in 
//  all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
// THE SOFTWARE.
////////////////////////////////////////////////////////////////////////////////
// #define v136 // uncomment this for v13.6 or v13.7
using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

#if !v136
// for 15.1 script approval so approval wizard knows this is a writeable script.
[assembly: ESAPIScript(IsWriteable = true)]
#endif

namespace VMS.TPS
{
    public class Script
    {
        // these three strings define the patient/study/image id for the image phantom that will be copied into the active patient.
        public static string QAPatientID_Trilogy = "Trilogy";
        public static string QAStudyID_Trilogy = "CT1";
        public static string QAImageID_Trilogy = "ArcCheck";

        public static string QAPatientID_iX = "iX";
        public static string QAStudyID_iX = "CT2";
        public static string QAImageID_iX = "MapCheck2";

        public static string QAMachine_Trilogy = "Trilogy";
        public static string QAMachine_iX = "iX";

        public static string QAId = "QA";

        public Script()
        {
        }

        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            Patient p = context.Patient;
            if (p == null)
                throw new ApplicationException("Please load a patient");

            ExternalPlanSetup plan = context.ExternalPlanSetup;
            if (plan == null)
                throw new ApplicationException("Please load an external beam plan that will be verified.");

            // Get or create course with Id 'QA'
            Course course = p.Courses.Where(o => o.Id == QAId).SingleOrDefault();
            if (course == null)
            {
                course = p.AddCourse();
                course.Id = QAId;
            }
            if (course.CompletedDateTime != null)
                MessageBox.Show("Course QA is set to 'COMPLETED', please set to 'ACTIVE'");

            p.BeginModifications();

            StructureSet ssQA = p.CopyImageFromOtherPatient(QAPatientID_Trilogy, QAStudyID_Trilogy, QAImageID_Trilogy);

            foreach (StructureSet ss in course.Patient.StructureSets)
            {
                if (plan.Beams.FirstOrDefault().TreatmentUnit.Name == QAMachine_Trilogy)
                {
                    if (ss.Image.Id == QAImageID_Trilogy)
                    {
                        ssQA = ss;
                        break;
                    }
                    else
                        ssQA = p.CopyImageFromOtherPatient(QAPatientID_Trilogy, QAStudyID_Trilogy, QAImageID_Trilogy);
                }
                else
                {
                    if (plan.Beams.FirstOrDefault().TreatmentUnit.Id == QAMachine_iX)
                    {
                        if (ss.Image.Id == QAImageID_iX)
                        {
                            ssQA = ss;
                            break;
                        }
                        else
                            ssQA = p.CopyImageFromOtherPatient(QAPatientID_iX, QAStudyID_iX, QAImageID_iX);
                    }
                    else
                    {
                        MessageBox.Show(string.Format("Treatment machine {0} in plan not recognized.", plan.Beams.FirstOrDefault().TreatmentUnit.Id));
                        ssQA = null;
                    }
                }
            }

#if false
        // Create an individual verification plan for each field.
        foreach (var beam in plan.Beams)
        {
            CreateVerificationPlan(course, new List<Beam> { beam }, plan, ssQA, beam.Id, calculateDose: false);
        }
#endif
            foreach (Beam beam in plan.Beams)
            {
                if (beam.ControlPoints.FirstOrDefault().PatientSupportAngle.ToString() != "0")
                {
                    planHasCouchKick = true;
                    MessageBox.Show("Plan has couch kick, please manually zero and recalculate/export.");
                    break;
                }                  
            }
            // Create a verification plan that contains all fields (Composite).
            ExternalPlanSetup verificationPlan = CreateVerificationPlan(course, plan.Beams, plan, ssQA, calculateDose: true);

            
            
             //navigate to verification plan
            //PlanSetup verifiedPlan = verificationPlan.VerifiedPlan;
           // if (plan != verifiedPlan)
           // {
            //    MessageBox.Show(string.Format("ERROR! verified plan {0} != loaded plan {1}", verifiedPlan.Id
           //         , plan.Id));
            //}
            MessageBox.Show(string.Format("Success - verification plan {0} created in course {1}.", verificationPlan.Id, course.Id));

        }
        /// <summary>
        /// Create verifications plans for a given treatment plan.
        /// </summary>
        public static ExternalPlanSetup CreateVerificationPlan(Course course, IEnumerable<Beam> beams, ExternalPlanSetup verifiedPlan, StructureSet verificationStructures,
                                                   bool calculateDose)
        {
            var verificationPlan = course.AddExternalPlanSetupAsVerificationPlan(verificationStructures, verifiedPlan);
            try
            {
                string verificationId = "";
                if (verifiedPlan.Id.Length >= 13)  //if over 13 chars in Id
                    verificationId = verifiedPlan.Id.Substring(0, (verifiedPlan.Id.Length - 2)) + "_A";
                else
                    verificationId = verifiedPlan.Id + "_A";
                verificationPlan.Id = verificationId;
            }
            catch (System.ArgumentException)
            {
                var message = "Plan already exists in QA course.";
                throw new Exception(message);
            }

            // Put isocenter to the center of the QAdevice
            VVector isocenter = verificationPlan.StructureSet.Image.UserOrigin;
            var beamList = verifiedPlan.Beams.ToList(); //used for looping later
            foreach (Beam beam in beams)
            {
                if (beam.IsSetupField)
                    continue;
                
                ExternalBeamMachineParameters MachineParameters =
                    new ExternalBeamMachineParameters(beam.TreatmentUnit.Id, beam.EnergyModeDisplayName, beam.DoseRate, beam.Technique.Id, string.Empty);

                if (beam.MLCPlanType.ToString() == "VMAT")
                {
                    // Create a new VMAT beam.
                    var collimatorAngle = beam.ControlPoints.First().CollimatorAngle;
                    var gantryAngleStart = beam.ControlPoints.First().GantryAngle;
                    var gantryAngleEnd = beam.ControlPoints.Last().GantryAngle;
                    var gantryDirection = beam.GantryDirection;
                    var metersetWeights = beam.ControlPoints.Select(cp => cp.MetersetWeight);
                    verificationPlan.AddVMATBeam(MachineParameters, metersetWeights, collimatorAngle, gantryAngleStart,
                        gantryAngleEnd, gantryDirection, 0.0, isocenter);
                    continue;                  
                }
                else
                {
                    if (beam.MLCPlanType.ToString() == "DoseDynamic")
                    {
                        // Create a new IMRT beam.
                        double gantryAngle;
                        double collimatorAngle;
                        if (beam.TreatmentUnit.Name == "Trilogy") //arccheck
                        {
                            gantryAngle = beam.ControlPoints.First().GantryAngle;
                            collimatorAngle = beam.ControlPoints.First().CollimatorAngle;
                        }

                        else //ix with only mapcheck
                        {
                            gantryAngle = 0.0;
                            collimatorAngle = 0.0;
                        }

                        var metersetWeights = beam.ControlPoints.Select(cp => cp.MetersetWeight);
                        verificationPlan.AddSlidingWindowBeam(MachineParameters, metersetWeights, collimatorAngle, gantryAngle,
                            0.0, isocenter);
                        continue;
                    }
                    else
                    {
                        var message = string.Format("Treatment field {0} is not VMAT or IMRT.", beam);
                        throw new Exception(message);
                    }
                }             
            }

            int i = 0;
            foreach (Beam verificationBeam in verificationPlan.Beams)
            {
                verificationBeam.Id = beamList[i].Id;
                i++;
            }
            
            foreach (Beam verificationBeam in verificationPlan.Beams)
            {
                foreach(Beam beam in verifiedPlan.Beams)
                {
                    if (verificationBeam.Id == beam.Id && verificationBeam.MLCPlanType.ToString() == "DoseDynamic")
                    {
                        var editableParams = beam.GetEditableParameters();
                        editableParams.Isocenter = verificationPlan.StructureSet.Image.UserOrigin;
                        verificationBeam.ApplyParameters(editableParams);                        
                        continue;
                    }
                }
            }
            // Set presciption
            const int numberOfFractions = 1;
            verificationPlan.SetPrescription(numberOfFractions, verifiedPlan.DosePerFraction, verifiedPlan.TreatmentPercentage);

            verificationPlan.SetCalculationModel(CalculationType.PhotonVolumeDose, verifiedPlan.GetCalculationModel(CalculationType.PhotonVolumeDose));

            CalculationResult res;
            if (verificationPlan.Beams.FirstOrDefault().MLCPlanType.ToString() == "DoseDynamic")
            {
                var getCollimatorAndGantryAngleFromBeam = beams.Count() > 1;
                var presetValues = (from beam in beams
                                    select new KeyValuePair<string, MetersetValue>(beam.Id, beam.Meterset)).ToList();
                res = verificationPlan.CalculateDoseWithPresetValues(presetValues);
            }
            else //vmat
                res = verificationPlan.CalculateDose();
            if (!res.Success)
            {
                var message = string.Format("Dose calculation failed for verification plan. Output:\n{0}", res);
               throw new Exception(message);
            }
            return verificationPlan;
        }
    }
}
