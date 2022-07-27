import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import axios from "axios";

export class ClientCredentialAuthenticationProvider implements AuthenticationProvider {
    public async getAccessToken(): Promise<string> {
        try {
            const url = process.env.APIURL || "https://yourapi.azurewebsites.net"
            let response = await axios.get(`${url}/graph`)
            if (response.status === 200) {
                return response.data;
            } else {
                throw new Error("Non 200OK response on obtaining token...")
            }
        }
        catch (error) {
            throw new Error("Error on obtaining token...")
        }
    }
    
}