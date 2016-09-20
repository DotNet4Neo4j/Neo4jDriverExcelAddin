namespace Neo4jDriverExcelAddin
{
    using System;

    internal class ExecuteCypherQueryArgs : EventArgs
    {
        public string Cypher { get; set; }
    }
}