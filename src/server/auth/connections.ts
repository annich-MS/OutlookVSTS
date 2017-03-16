import * as Knex from "knex";

interface IConnection {
  // | object is dumb, but azure's config is not implemented in knex
  [endpoint: string]: Knex.Config | object;
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
      options: {
        database: process.env.DB_NAME,
        encrypt: true,
      },
      password: process.env.DB_PASSWORD,
      server: process.env.DB_SERVER,
      user: process.env.DB_USER,
    },
    migrations: {
      directory: `${__dirname}/migrations`,
    },
    pool: {
      max: 10,
      min: 2,
    },
  },

};

export default connections;
