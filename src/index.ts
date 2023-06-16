import { google } from 'googleapis';
import { JWT } from 'google-auth-library';

class GoogleSpreadSheetsApis {

    jwt: JWT | null = null;
    client_email: string = '';
    private_key: string = '';


    constructor(client_email: string, private_key: string) {
        // TODO import { client_email, private_key } from '*.json';
        this.client_email = client_email;
        this.private_key = private_key;
    }

    authorize() {
        this.jwt = new google.auth.JWT(this.client_email, undefined, this.private_key, ['https://www.googleapis.com/auth/spreadsheets']);
    }

    append(spreadsheetId: string, sheetName: string, row: any[]) {
        if (this.jwt === null) {
            throw new Error('no jwt');
        }

        const sheets = google.sheets({ version: 'v4', auth: this.jwt });

        return sheets.spreadsheets.values.append({
            spreadsheetId,
            valueInputOption: 'USER_ENTERED',
            range: `${sheetName}!A1`,
            requestBody: {
                values: [row]
            }
        });
    }

    batchUpdate(spreadsheetId: string, sheetName: string, data: any[][]) {
        if (this.jwt === null) {
            throw new Error('no jwt');
        }

        const sheets = google.sheets({ version: 'v4', auth: this.jwt });

        return sheets.spreadsheets.values.batchUpdate({
            spreadsheetId,
            requestBody: {
                valueInputOption: 'USER_ENTERED',
                data: [
                    {
                        range: `${sheetName}!A1`,
                        values: data
                    }
                ]
            }
        });
    }
}

export default GoogleSpreadSheetsApis;
