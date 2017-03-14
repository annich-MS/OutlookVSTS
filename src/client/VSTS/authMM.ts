import { Rest } from '../RestHelpers/rest';

/**
 * interface for callback
 * @interface IAuthStateCallback 
 */
export interface IAuthStateCallback { (state: string): void; }

/**
 * Connects to user database
 * @class Auth
 */
export class Auth {
    /**
     * check user database for token associated with user email
     * @param {string} user - user's email address
     * @param {IAuthStateCallback} callback
     * @return {void}
     */
    public static getAuthState(callback: IAuthStateCallback): void {
        Rest.getUser((user: string) => {
            $.get('./authenticate/db?user=' + user + '&trash=' + (Math.random() * 1000), (output) => {
                console.log(output);
                if (output === 'success') {
                    callback('success');
                } else {
                    callback('failure');
                }
            });
        });
    }
}
