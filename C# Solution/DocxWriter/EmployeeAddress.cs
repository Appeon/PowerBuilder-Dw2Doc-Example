using System.Text.Json.Serialization;

namespace Appeon.DotnetDemo.DocumentWriter
{
    public class EmployeeAddress
    {
        [JsonPropertyName("BusinessEntityID")]
        public long BusinessEntityId { get; set; }

        [JsonPropertyName("AddressLine1")]
        public string? AddressLine1 { get; set; }

        [JsonPropertyName("AddressLine2")]
        public string? AddressLine2 { get; set; }

        [JsonPropertyName("City")]
        public string? City { get; set; }

        [JsonPropertyName("PostalCode")]
        public string? PostalCode { get; set; }

        [JsonPropertyName("StateProvinceCode")]
        public string? StateProvinceCode { get; set; }

        [JsonPropertyName("CountryRegionCode")]
        public string? CountryRegionCode { get; set; }

        [JsonPropertyName("Name")]
        public string? StateName { get; set; }
    }
}
