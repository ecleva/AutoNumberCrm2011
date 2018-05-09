using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Crm.Sdk;
using Microsoft.Xrm;
using System.Diagnostics;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace AutoNumberCrm2011
{
    public class AutoNumber : IPlugin
    {

        #region IPlugin Members

        public void Execute(IServiceProvider serviceProvider)
        {
            string mode = string.Empty;
            int stage = 0;
            string errorMessage = string.Empty;
            string entityName = string.Empty;

            int counter = 0;
            string prefix = string.Empty;
            string suffix = string.Empty;
            string autoNumberAttribute = string.Empty;

            


               ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

                  //Get the Event Name
            mode = context.MessageName;
            //Get the Stage (Pre/Post)
            stage = context.Stage;
            //Get the Fired Entity Name
            entityName = context.PrimaryEntityName;
            
            
            //Fired on Pre-Create Mode
            //if (mode.Equals("Create"))
            {  
            
            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {               

                Entity firedEntity = (Entity)context.InputParameters["Target"];
               
                    try
                    {
                        //Create the object of ICrmService
                      //  ICrmService oService = (ICrmService)context.CreateCrmService(true);
                        // Obtain the organization service reference.
                        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                        IOrganizationService service = serviceFactory.CreateOrganizationService(null);
                                                
                        //Retrieve AutoNumber Config values using Dynamic Entity
                       // DynamicEntity dynAutoNumber = (DynamicEntity)crm_helper.GetAutoNumberConfig(oService, "pnp_autonumber", entityName);
                        Entity dynAutoNumber = (Entity)crm_helper.GetAutoNumberConfig(service, "pnp_autonumber", entityName);
                                 
                        if (dynAutoNumber.Contains("pnp_counter"))
                            counter = Convert.ToInt32(dynAutoNumber.Attributes["pnp_counter"].ToString());

                        if (dynAutoNumber.Contains("pnp_prefix"))
                            prefix = dynAutoNumber.Attributes["pnp_prefix"].ToString();

                        if (dynAutoNumber.Contains("pnp_suffix"))
                            suffix = dynAutoNumber.Attributes["pnp_suffix"].ToString();

                        if (dynAutoNumber.Contains("pnp_autonumberattribute"))
                            autoNumberAttribute = dynAutoNumber.Attributes["pnp_autonumberattribute"].ToString();

                        if (entityName.Equals("pnp_environmentalliability"))
                        {
                            int liabilityCounter = Convert.ToInt16(suffix);

                            UpdateAutoNumberForLiabilityCase(service, dynAutoNumber, firedEntity, liabilityCounter);

                            //Create variable to store autonumber value
                            String updateAutoNumber = String.Empty;
                            //Set the autonumber value to current year
                            String currYear = DateTime.Now.Year.ToString().Substring(2);
                            if (currYear.Length < 2)
                                currYear = "0" + currYear;

                            if (liabilityCounter < 10)
                            {
                                updateAutoNumber = "00" + liabilityCounter.ToString();
                            }
                            if (liabilityCounter > 9 && liabilityCounter < 100)
                            {
                                updateAutoNumber = "0" + liabilityCounter.ToString();
                            }
                            if (liabilityCounter > 99)
                            {
                                updateAutoNumber = liabilityCounter.ToString();
                            }

                            //Update the Fired Entity with the AutoNumber
                            updateAutoNumber = prefix + currYear + updateAutoNumber;

                            //If required attribute exist in property collection then update it or else create new property
                            if (firedEntity.Contains(autoNumberAttribute))
                            {
                                firedEntity.Attributes[autoNumberAttribute] = updateAutoNumber;
                            }
                            else
                            {
                                firedEntity.Attributes.Add(autoNumberAttribute, updateAutoNumber);
                            }

                        }
                        else
                        {
                            string updateAutoNumber = string.Empty;

                            //daloor 29May2014 need 4 digits for article 27
                            //modified by bhavik shah on 22-jan-2018

                            if (entityName.Equals("pnp_article27") && mode.ToUpper().Equals("UPDATE")
                                && firedEntity.Contains("statuscode") && ((OptionSetValue)firedEntity["statuscode"]).Value == 125560000 && !firedEntity.Contains(updateAutoNumber))
                            {
                                //Update autonumber counter in custom entity
                                UpdateAutoNumber(service, dynAutoNumber, firedEntity, counter);

                                //Create variable to store autonumber value
                                
                                updateAutoNumber = counter.ToString();

                                
                                Entity Article11 = service.Retrieve("pnp_article27", firedEntity.Id, new ColumnSet(autoNumberAttribute));
                                //updateAutoNumber = prefix + counter.ToString() + suffix;

                                //added on 22-feb-2018
                                updateAutoNumber = prefix + FormatAutoNumber(counter.ToString()) + suffix;


                                if (Article11.Contains(autoNumberAttribute) && !String.IsNullOrWhiteSpace(Article11[autoNumberAttribute].ToString()))
                                {
                                    //DO Nothing- Not overrite AN
                                }
                                else
                                {
                                    firedEntity.Attributes[autoNumberAttribute] = updateAutoNumber;
                                }

                                //updateAutoNumber = prefix + FormatAutoNumber(counter.ToString()) + suffix;
                                ////If required attribute exist in property collection then update it or else create new property
                                //if (firedEntity.Contains(autoNumberAttribute))
                                //{
                                //    firedEntity.Attributes[autoNumberAttribute] = updateAutoNumber;
                                //}
                                //else
                                //{
                                //    firedEntity.Attributes.Add(autoNumberAttribute, updateAutoNumber);
                                //}
                            }
                            //Only for Article 11, Auto Number works on Update of StatusCode to Recieved
                            //modified by bhavik shah on 03-sep-2017
                            else if (entityName.Equals("pnp_article11request") && mode.ToUpper() == "UPDATE"
                                && firedEntity.Contains("statuscode") && ((OptionSetValue)firedEntity["statuscode"]).Value == 3 && !firedEntity.Contains(updateAutoNumber))
                            {

                                //Update autonumber counter in custom entity
                                UpdateAutoNumber(service, dynAutoNumber, firedEntity, counter);

                                //Create variable to store autonumber value

                                updateAutoNumber = counter.ToString();

                                //Read AN field value

                                Entity Article11 = service.Retrieve("pnp_article11request", firedEntity.Id, new ColumnSet(autoNumberAttribute));
                                //updateAutoNumber = prefix + counter.ToString() + suffix;

                                //added on 22-feb-2018
                                updateAutoNumber = prefix + FormatAutoNumber(counter.ToString()) + suffix;

                                if (Article11.Contains(autoNumberAttribute) && !String.IsNullOrWhiteSpace(Article11[autoNumberAttribute].ToString()))
                                {
                                    //DO Nothing- Not overrite AN
                                }
                                else
                                {
                                    firedEntity.Attributes[autoNumberAttribute] = updateAutoNumber;
                                }

                                //If required attribute exist in property collection then update it or else create new property
                                //if (firedEntity.Contains(autoNumberAttribute))
                                //{
                                //    firedEntity.Attributes[autoNumberAttribute] = updateAutoNumber;
                                //}
                                //else
                                //{
                                //    firedEntity.Attributes.Add(autoNumberAttribute, updateAutoNumber);
                                //}
                            }
                            //else if (entityName != "pnp_article27")
                            else if (entityName != "pnp_article27" && entityName != "pnp_article11request")
                            {
                                //Update autonumber counter in custom entity
                                UpdateAutoNumber(service, dynAutoNumber, firedEntity, counter);

                                //Create variable to store autonumber value

                                updateAutoNumber = counter.ToString();

                                //updateAutoNumber = prefix + counter.ToString() + suffix;
                                //added on 22-feb-2018
                                updateAutoNumber = prefix + FormatAutoNumber(counter.ToString()) + suffix;

                                //If required attribute exist in property collection then update it or else create new property
                                if (firedEntity.Contains(autoNumberAttribute))
                                {
                                    firedEntity.Attributes[autoNumberAttribute] = updateAutoNumber;
                                }
                                else
                                {
                                    firedEntity.Attributes.Add(autoNumberAttribute, updateAutoNumber);
                                }
                            }
                            //daloor 29May2014 need 4 digits for article 27                     



                           

                        }

                        log.logInfo("AutoNumber Plugin logic completed");

                    }

                    catch (SystemException ex)
                    {
                        throw new InvalidPluginExecutionException("Sorry !!! Record cannot be saved. Error with AutoNumber");
                    }

                }
            }
        }

        #endregion

        public string FormatAutoNumber(string autonumber)
        {
            string retValue = string.Empty;

            try
            {
                switch (autonumber.Length)
                {
                    case 1:
                        retValue = "000" + autonumber;
                        break;
                    case 2:
                        retValue = "00" + autonumber;
                        break;
                    case 3:
                        retValue = "0" + autonumber;
                        break;
                    default:
                        retValue = autonumber;
                        break;
                }

            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                throw ex;
            }
            catch (SystemException ex)
            {
                throw ex;
            }

            return retValue;
        }


        public void UpdateAutoNumber(IOrganizationService service, Entity dynAutoNumber, Entity firedEntity, int counter)
        {
            try
            {
                //Increment and Update the AutoNumber Entity
                int newCounter = counter + 1;

                dynAutoNumber.Attributes["pnp_counter"] = newCounter;

                UpdateRequest updatereq = new UpdateRequest();
                updatereq.Target = dynAutoNumber;
                UpdateResponse updateresponse = (UpdateResponse)service.Execute(updatereq);

              

            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                throw ex;
            }
            catch (SystemException ex)
            {
                throw ex;
            }

        }


        public void UpdateAutoNumberForLiabilityCase(IOrganizationService service, Entity dynAutoNumber, Entity firedEntity, int counter)
        {
            try
            {
                //Increment and Update the AutoNumber Entity
                int newCounter = counter + 1;

                dynAutoNumber.Attributes["pnp_suffix"] = newCounter.ToString();
                
                UpdateRequest updatereq = new UpdateRequest();
                updatereq.Target = dynAutoNumber;
                UpdateResponse updateresponse = (UpdateResponse)service.Execute(updatereq);

            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                throw ex;
            }
            catch (SystemException ex)
            {
                throw ex;
            }

        }
       
    }

   
}
