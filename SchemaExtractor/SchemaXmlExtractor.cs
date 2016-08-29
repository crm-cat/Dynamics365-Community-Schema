using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Dyn365CommunitySchema
{
    /// <summary>
    /// Extracts the schema for the target organization into xml files for easy comparison in source control
    /// </summary>
    public class SchemaXmlExtractor
    {
        private CrmServiceClient _service;
        private string[] _entityLogicalNames;
        private string[] _systemEntitiesToIgnore;
        private string _rootFolderPath;
        private string _targetFolder;
        private string[] _ignoreAttributesStartingWith;
        private string[] _ignoreAttributesEndingWith;
        private string[] _attributeElementsToStrip;
        private string _connectionString;
        private Dictionary<string, AttributeMetadata> _lookupAttributes;

        public SchemaXmlExtractor(string[] entityLogicalNames, string connectionString, string targetFolder)
        {
            _entityLogicalNames = entityLogicalNames;
            _connectionString = connectionString;
            if (targetFolder == null)
            {
                _targetFolder = AppDomain.CurrentDomain.BaseDirectory;
            }
            else
            {
                _targetFolder = targetFolder;
            }

            _entityLogicalNames = new string[] {
                "account",
                "contact",
                "customeraddress"};

            _systemEntitiesToIgnore = new string[] {
                "sharepointdocument",
                "sharepointdocumentlocation",
                "activitypointer",
                "activityparty",
                "listmember",
                "socialactivity",
                "socialprofile",
                "syncerror",
                "bulkoperationlog",
                "slakpiinstance",
                "userentityinstancedata",
                "entitlement",
                "bookableresource",
                "annotation",
                "bulkdeletefailure",
                "customerrelationship",
                "duplicaterecord",
                "asyncoperation",
                "processsession",
                "postrole",
                "postregarding",
                "postfollow",
                "mailboxtrackingfolder",
                "connection",
                "customeropportunityrole",
                "principalobjectattributeaccess",
                "imagedescriptor",
                "sla",
                "equipment",
                "listmember",
                "subscription"
            };

            _ignoreAttributesStartingWith = new string[] {
                "createdon",
                "createdby",
                "createdbyexternalparty",
                "modifiedon",
                "modifiedby",
                "modifiedonbehalfby",
                "overriddencreatedon",
                "ownerid",
                "utcconversiontimezonecode",
                "yominame",
                "versionnumber",
                "importsequencenumber",
                "masterid",
                "merged",
                "timezoneruleversionnumber",
                "traversedpath",
                "stageid",
                "processid",
                "participatesinworkflow",
                "masterid",
                "businessunitid",
                "owningteam",
                "owninguser",
                "ms_traceid",
                "yomi",
                "utcoffeset"
                };

            _ignoreAttributesEndingWith = new string[] {
                "_base",
                "_date",
                "_state" };

            _attributeElementsToStrip = new string[] {
                "UserLocalizedLabel",
                "IsAuditEnabled",
                "MetadataId",
                "ColumnNumber",
                "IsGlobalFilterEnabled",
                "IsManaged",
                "IsRenameable",
                "IsRetrievable",
                "IsSearchable",
                "IsSecured",
                "IsSortableEnabled",
                "IsValidForAdvancedFind",
                "IsValidForCreate",
                "IsValidForRead",
                "IsValidForUpdate",
                "CanBeSecuredForCreate",
                "CanBeSecuredForRead",
                "CanBeSecuredForUpdate",
                "CanModifyAdditionalSettings",
                "IsCustomizable",
                "IntroducedVersion",
                "DeprecatedVersion",
                "HasChanged"
            };
        }

        public void Connect()
        {    
            _service = new CrmServiceClient(_connectionString);
            _rootFolderPath = GetFolder(_targetFolder, "dyn365-community-schema");
            var response = (WhoAmIResponse)_service.Execute(new WhoAmIRequest());
        }

        public void Extract()
        {
            RetrieveMetadataChangesResponse metadataRequest = GetEntityMetadata();
            SerialiseEntities(metadataRequest);
        }

        private void SerialiseEntities(RetrieveMetadataChangesResponse metadataRequest)
        {
            var serialiser = new DataContractSerializer(typeof(AttributeMetadata));

            // Serialise each attribute and relationship
            foreach (var entity in metadataRequest.EntityMetadata)
            {
                var entityFolder = GetFolder(_rootFolderPath, entity.LogicalName);
                _lookupAttributes = new Dictionary<string, AttributeMetadata>();

                foreach (var attribute in entity.Attributes)
                {
                    if (attribute.DeprecatedVersion != null)
                        continue;

                    if (attribute.IsLogical == true)
                        continue;

                    if (IsAttributeIgnored(attribute.LogicalName))
                        continue;

                    if (attribute.AttributeType == AttributeTypeCode.Lookup)
                    {
                        // Don't output here - output as part of the relationship metadata
                        _lookupAttributes.Add(attribute.LogicalName, attribute);
                        continue;
                    }

                    var xmlOutput = SerialiseToStringBuilder(serialiser, attribute);
                    var attributeFilePath = Path.Combine(entityFolder, attribute.LogicalName + ".xml");
                    StripAndWriteXml(xmlOutput, attributeFilePath);
                }

                // Output the relationship xml
                SerialiseRelationships(entity, entityFolder);
            }
        }

        private void SerialiseRelationships(EntityMetadata entity, string entityFolder)
        {
            var manyToOneFolder = GetFolder(entityFolder, "many-to-one");
            if (entity.ManyToOneRelationships != null)
            {
                SerialiseOneToManyRelationship(entity.ManyToOneRelationships, manyToOneFolder);
            }

            var manyToManyFolder = GetFolder(entityFolder, "many-to-many");
            var manyToManySerialiser = new DataContractSerializer(typeof(ManyToManyRelationshipMetadata));
            if (entity.ManyToManyRelationships != null)
            {
                foreach (var relationship in entity.ManyToManyRelationships)
                {
                    if (_systemEntitiesToIgnore.Contains(relationship.Entity1LogicalName))
                        continue;

                    if (_systemEntitiesToIgnore.Contains(relationship.Entity2LogicalName))
                        continue;

                    var relationshipFilePath = Path.Combine(manyToManyFolder, relationship.SchemaName + ".xml");
                    var xmlOutput = SerialiseToStringBuilder(manyToManySerialiser, relationship);
                    StripAndWriteXml(xmlOutput, relationshipFilePath);
                }
            }
        }

        private RetrieveMetadataChangesResponse GetEntityMetadata()
        {
            var entityFilter = new MetadataFilterExpression(LogicalOperator.And);

            entityFilter.Conditions.Add(
                new MetadataConditionExpression(
                    "LogicalName",
                    MetadataConditionOperator.In,
                    _entityLogicalNames));

            var entityProperties = new MetadataPropertiesExpression()
            {
                AllProperties = false
            };

            entityProperties.PropertyNames.AddRange(new string[] {
                "Attributes",
                "ManyToManyRelationships",
                "OneToManyRelationships",
                "ManyToOneRelationships"});

            // A filter expression to apply the optionsetAttributeTypes condition expression
            var attributeFilter = new MetadataFilterExpression(LogicalOperator.Or);
            
            // A Properties expression to limit the properties to be included with attributes
            var attributeProperties = new MetadataPropertiesExpression() { AllProperties = true };

            // A label query expression to limit the labels returned to only those for the output preferred language
            var labelQuery = new LabelQueryExpression();
            labelQuery.FilterLanguages.Add(1033);

            // An entity query expression to combine the filter expressions and property expressions for the query.
            var entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter,
                Properties = entityProperties,
                AttributeQuery = new AttributeQueryExpression()
                {
                    Criteria = attributeFilter,
                    Properties = attributeProperties
                },
                LabelQuery = labelQuery
            };

            // Retrieve the metadata for the query without a ClientVersionStamp
            var request = new RetrieveMetadataChangesRequest()
            {
                ClientVersionStamp = null,
                Query = entityQueryExpression
            };

            var metadataRequest = (RetrieveMetadataChangesResponse)_service.Execute(request);
            return metadataRequest;
        }

        private void SerialiseOneToManyRelationship(OneToManyRelationshipMetadata[] relationships, string folder)
        {
            var oneToManySerialiser = new DataContractSerializer(typeof(ManyToOneRelationship));

            foreach (var relationship in relationships)
            {
                // Don't include system relationships
                if (relationship.ReferencingAttribute == "regardingobjectid")
                    continue;

                if (_systemEntitiesToIgnore.Contains(relationship.ReferencingEntity))
                    continue;

                if (IsAttributeIgnored(relationship.ReferencingAttribute))
                    continue;

                if (_systemEntitiesToIgnore.Contains(relationship.ReferencedEntity))
                    continue;

                if (IsAttributeIgnored(relationship.ReferencedAttribute))
                    continue;

                // Get the corresponding lookup attribute on this entity if there is one
                AttributeMetadata lookupAttribute = null;
                if (_lookupAttributes.ContainsKey(relationship.ReferencingAttribute))
                {
                    lookupAttribute = _lookupAttributes[relationship.ReferencingAttribute];
                }

                var relationshipExtended = new ManyToOneRelationship
                {
                    Relationship = relationship,
                    LookupAttribute = lookupAttribute
                };

                var relationshipFilePath = Path.Combine(folder, relationship.SchemaName + ".xml");
                var xmlOutput = SerialiseToStringBuilder(oneToManySerialiser, relationshipExtended);
                StripAndWriteXml(xmlOutput, relationshipFilePath);
            }
        }

        private bool IsAttributeIgnored(string attribute)
        {
            var ignoreAttribute = false;

            // Don't include some system attributes (starting with)
            ignoreAttribute = (_ignoreAttributesStartingWith.Where(a => attribute.StartsWith(a)).FirstOrDefault() != null);

            // Don't include some system attributes (ending with)
            ignoreAttribute = ignoreAttribute || (_ignoreAttributesEndingWith.Where(a => attribute.EndsWith(a)).FirstOrDefault() != null);

            return ignoreAttribute;
        }

        private static StringBuilder SerialiseToStringBuilder(DataContractSerializer serialiser, object value)
        {
            StringBuilder xmlOutput = new StringBuilder();
            using (var xmlWriter = XmlWriter.Create(xmlOutput))
            {
                // Serialise
                serialiser.WriteObject(xmlWriter, value);
            }
            return xmlOutput;
        }

        private void StripAndWriteXml(StringBuilder xmlOutput, string xmlFilePath)
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlOutput.ToString());

            // Remove the unneeded elements as they just clutter the xml
            foreach (var elementName in _attributeElementsToStrip)
            {
                RemoveNodes(xml, elementName);
            }

            using (var writer = XmlWriter.Create(xmlFilePath, new XmlWriterSettings {
                Indent = true,
                IndentChars = "\t",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = true }))
            {
                xml.WriteContentTo(writer);
            }
        }

        /// <summary>
        /// Remove a set of element nodes from an xml document
        /// </summary>
        /// <param name="xml">Document to remove nodes from</param>
        /// <param name="removeAttribute">Element name to remove</param>
        private static void RemoveNodes(XmlDocument xml, string removeAttribute)
        {
            var nodesToRemove = xml.SelectNodes("//*[local-name()='" + removeAttribute + "']");
            foreach (XmlNode node in nodesToRemove)
            {
                node.ParentNode.RemoveChild(node);
            }
        }

        /// <summary>
        /// Get a folder and create it if it doesn't exist
        /// </summary>
        /// <param name="rootFolder">Path to get the folder in</param>
        /// <param name="folderName">Name of the folder</param>
        /// <returns></returns>
        private static string GetFolder(string rootFolder, string folderName)
        {
            var oneToManyFolder = Path.Combine(rootFolder, folderName);
            if (!Directory.Exists(oneToManyFolder))
            {
                Directory.CreateDirectory(oneToManyFolder);
            }
            return oneToManyFolder;
        }
    }
}
