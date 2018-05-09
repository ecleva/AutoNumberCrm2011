using System;
using System.Configuration;
using System.Collections;
using System.Globalization;
using System.Xml;
using System.Text;
using System.Reflection;
using Microsoft.Crm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
//using Microsoft.Crm.SdkTypeProxy;
//using Microsoft.Crm.Sdk.Query ;
//using Microsoft.Crm.Sdk.Metadata;
//using Microsoft.Crm.SdkTypeProxy.Metadata;
using System.Web.Services.Protocols;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

//***************************************************
//© 2008  All rights reserved.
//***************************************************
// Date		    Who			    Description
// 15-Aug-2008	Denny A1oor		AutoNumber Creation
//***************************************************

namespace AutoNumberCrm2011
{
	/// <summary>
	/// Summary description for crm_helper.
	/// </summary>Michael
	public class crm_helper
	{
		    
        
        public static Entity GetAutoNumberConfig(IOrganizationService oService, string strObjectType, string entityName)
        {
            EntityCollection ec = null;

            try
            {
                QueryExpression query = new QueryExpression();

                ColumnSet qpColSet = new ColumnSet();
                string[] strColumnList = { "pnp_autonumberid", "pnp_prefix", "pnp_counter", "pnp_suffix", "pnp_autonumberattribute" };
                qpColSet.AddColumns(strColumnList);
                query.ColumnSet = qpColSet;
                query.EntityName = strObjectType;

                ConditionExpression ce = new ConditionExpression("pnp_name", ConditionOperator.Equal, entityName.Trim());
                query.Criteria.AddCondition(ce);                

                RetrieveMultipleRequest retrieve = new RetrieveMultipleRequest();
                retrieve.Query = query;

                ec = oService.RetrieveMultiple(query);                
            }
            catch (SoapException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return (Entity)ec.Entities[0];
        }


	}
}
