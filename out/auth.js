var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { inTeams, ensureTeamsSdkInitialized, getAuthTokenFromTeams } from "./teams.js";
var msalConfig = {
    auth: {
        clientId: '0948fd6f-c5b0-460f-b59a-3481b7dd4702',
        authority: 'https://login.microsoftonline.com/6c51c659-9d52-41af-81f7-dde16380e813',
        redirectUri: "https://mkirpichnik.github.io/AI-Platform",
        postLogoutRedirectUri: "https://mkirpichnik.github.io/AI-Platform"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false // Set this to "true" if you are having issues on IE11 or Edge
    }
};
// MSAL request object to use over and over
var msalRequest = {
    scopes: ["api://0948fd6f-c5b0-460f-b59a-3481b7dd4702/authorized_user"]
};
var msalClient = new msal.PublicClientApplication(msalConfig);
var getAppAccessTokenPromise; // Cache the promise so we only do the work once on this page
export function getAppAccessToken() {
    if (!getAppAccessTokenPromise) {
        getAppAccessTokenPromise = getAppAccessToken2();
    }
    return getAppAccessTokenPromise;
}
// Here we do the work to log the user in and get the employee ID
function getAppAccessToken2() {
    return __awaiter(this, void 0, void 0, function () {
        var accessToken;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getAccessToken()];
                case 1:
                    accessToken = _a.sent();
                    console.log(accessToken);
                    return [2 /*return*/, accessToken];
            }
        });
    });
}
var getAccessTokenPromise; // Cache the promise so we only do the work once on this page
export function getAccessToken() {
    if (!getAccessTokenPromise) {
        getAccessTokenPromise = getAccessToken2();
    }
    return getAccessTokenPromise;
}
function getAccessToken2() {
    return __awaiter(this, void 0, void 0, function () {
        var error_1, accounts, accessToken, tokenResponse, error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, inTeams()];
                case 1:
                    if (!_a.sent()) return [3 /*break*/, 4];
                    return [4 /*yield*/, ensureTeamsSdkInitialized()];
                case 2:
                    _a.sent();
                    return [4 /*yield*/, getAuthTokenFromTeams()];
                case 3: return [2 /*return*/, _a.sent()];
                case 4: 
                // If we were waiting for a redirect with an auth code, handle it here
                return [4 /*yield*/, msalClient.handleRedirectPromise()];
                case 5:
                    // If we were waiting for a redirect with an auth code, handle it here
                    _a.sent();
                    _a.label = 6;
                case 6:
                    _a.trys.push([6, 8, , 10]);
                    return [4 /*yield*/, msalClient.ssoSilent(msalRequest)];
                case 7:
                    _a.sent();
                    return [3 /*break*/, 10];
                case 8:
                    error_1 = _a.sent();
                    console.log(error_1);
                    return [4 /*yield*/, msalClient.loginRedirect(msalRequest)];
                case 9:
                    _a.sent();
                    return [3 /*break*/, 10];
                case 10:
                    accounts = msalClient.getAllAccounts();
                    if (accounts.length === 1) {
                        debugger;
                        console.log(accounts);
                        // msalRequest.account = accounts[0];
                    }
                    else {
                        throw ("Error: Too many or no accounts logged in");
                    }
                    _a.label = 11;
                case 11:
                    _a.trys.push([11, 13, , 14]);
                    return [4 /*yield*/, msalClient.acquireTokenSilent(msalRequest)];
                case 12:
                    tokenResponse = _a.sent();
                    accessToken = tokenResponse.accessToken;
                    return [2 /*return*/, accessToken];
                case 13:
                    error_2 = _a.sent();
                    if (error_2 instanceof msal.InteractionRequiredAuthError) {
                        console.warn("Silent token acquisition failed; acquiring token using redirect");
                        msalClient.acquireTokenRedirect(msalRequest);
                    }
                    else {
                        throw (error_2.message);
                    }
                    return [3 /*break*/, 14];
                case 14: return [2 /*return*/];
            }
        });
    });
}
// Headers for use in Fetch (HTTP) requests when calling anonymous web services
// in the server side of this app.
export function getFetchHeadersAnon() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, ({
                    "content-type": "application/json"
                })];
        });
    });
}
// Headers for use in Fetch (HTTP) requests when calling authenticated web services
// in the server side of this app. Authentication is sent in a cookie, so no
// additional headers are required.
// Other implementations of this module may insert an Authorization header here
export function getFetchHeadersAuth() {
    return __awaiter(this, void 0, void 0, function () {
        var accessToken;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getAccessToken()];
                case 1:
                    accessToken = _a.sent();
                    return [2 /*return*/, ({
                            "content-type": "application/json",
                            "authorization": "Bearer ".concat(accessToken)
                        })];
            }
        });
    });
}
//# sourceMappingURL=auth.js.map