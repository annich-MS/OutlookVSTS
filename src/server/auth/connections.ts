import * as Knex from "knex";

interface IConnection {
  [endpoint: string]: Knex.Config;
}

const connections: IConnection = {

  development: {
    client: "sqlite3",
    connection: {
      filename: "./dev.sqlite3",
    },
    migrations: {
      directory: `${__dirname}/migrations`,
    },
    useNullAsDefault: true,
  },

    production: {
    client: "mssql",
    connection: {
      database: process.env.DB_NAME,
      host: process.env.DB_SERVER,
      password: process.env.DB_PASSWORD,
      user: process.env.DB_USER,
    },
    migrations: {
      directory: __dirname + "bin/db/migrations",
    },
    pool: {
      max: 10,
      min: 2,
    },
  },

};

export default connections;
