import * as Promise from "bluebird";
import * as Knex from "knex";
import Token from "../token";

export function up(knex: Knex): Promise<void[]> {
    return Promise.all([
        knex.schema.createTable(Token.TableName, (t: Knex.CreateTableBuilder) => {
            t.string(Token.IdKey).primary();
            t.string(Token.TokenKey, 1000);
            t.dateTime(Token.ExpiryKey);
            t.string(Token.RefreshKey, 1000);
        }),
    ]);
};

export function down (knex: Knex): Promise<void[]> {
    return Promise.all([
        knex.schema.dropTable(Token.TableName),
    ]);
};
