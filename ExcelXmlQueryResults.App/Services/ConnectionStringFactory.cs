using ExcelXmlQueryResults.App.Models;
using Microsoft.Data.SqlClient;

namespace ExcelXmlQueryResults.App.Services;

public sealed class ConnectionStringFactory
{
    public string Build(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        var builder = new SqlConnectionStringBuilder
        {
            DataSource = settings.Server,
            InitialCatalog = settings.Database,
            IntegratedSecurity = settings.ConnectionMode == ConnectionMode.IntegratedSecurity,
        };

        if (settings.ConnectionMode == ConnectionMode.UsernamePassword)
        {
            builder.UserID = settings.Username;
            builder.Password = settings.Password;
        }

        return builder.ConnectionString;
    }
}
