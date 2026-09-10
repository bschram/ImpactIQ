[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "Sales_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared Sales = let
    Source = Sql.Database("sql-finance.contoso.gov", "Ledger"),
    T = Source{[Schema = "dbo", Item = "Sales"]}[Data]
in
    T;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "Customers_DataDestination"]]}]
shared #"Customers 2024" = let
    Source = Sales
in
    Source;
shared Sales_DataDestination = let Pattern = Lakehouse.Contents([HierarchicalNavigation = null]) in Pattern;
