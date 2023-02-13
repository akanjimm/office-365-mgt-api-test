const config = {
  authorityHost: process.env.M365_AUTHORITY_HOST,
  tenantId: process.env.M365_TENANT_ID,
  clientId: process.env.M365_CLIENT_ID,
  clientSecret: process.env.M365_CLIENT_SECRET,

  /* Uncomment the lines below when using an sql server database
     Add azure sql database feature using teams toolkit extension if you are using
     azure sql database offered by teams toolkit
     
  sqlEndpoint: process.env.SQL_ENDPOINT,
  sqlUsername: process.env.SQL_USER_NAME,
  sqlPassword: process.env.SQL_PASSWORD,
  sqlDatabaseName: process.env.SQL_DATABASE_NAME,
  identityID: process.env.IDENTITY_ID // this is specific to Azure sql db
  */

  /* Add the following environment variables to the .env.teamsfx.local file and remove it from here
  SQL_ENDPOINT_OLD_SUBSCRIPTION=<value>
  SQL_USER_NAME_OLD_SUBSCRIPTION=<value>
  SQL_PASSWORD_OLD_SUBSCRIPTION=<value>
  SQL_DATABASE_NAME_OLD_SUBSCRIPTION=<value>
  IDENTITY_ID_OLD_SUBSCRIPTION=<value>
  */
};

module.exports = config;
