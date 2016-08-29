using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Dyn365CommunitySchema
{
    [DataContract]
    public class ManyToOneRelationship
    {
        [DataMember]
        public AttributeMetadata LookupAttribute { get; set; }
        [DataMember]
        public OneToManyRelationshipMetadata Relationship {get; set;}
    }
}
