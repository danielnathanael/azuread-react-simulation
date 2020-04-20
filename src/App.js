import React from 'react';
import './App.css'
import { loginRequest, msal, graphConfig } from './azure-active-directory'

class App extends React.Component {
  constructor(props) {
    super()

    this.state = {
      idToken: localStorage.getItem('idToken'),
      accessToken: localStorage.getItem('accessToken'),
      name: '',
      username: '',
      firstName: '',
      lastName: '',
      businessPhones: '',
      jobTitle: '',
      mail: '',
      mobilePhone: '',
      officeLocation: '',
      preferredLanguage: '',
      id: ''
    }
  }

  signIn() {
    msal.loginPopup(loginRequest)
      .then(response => {
        localStorage.setItem('idToken', response.idToken.rawIdToken)
        this.setState({
          idToken: response.idToken.rawIdToken
        })
      }).catch(error => {
        console.log(error);
      });
  }

  getAccessToken() {
    return msal.acquireTokenSilent(loginRequest)
      .then(response => {
        localStorage.setItem('accessToken', response.accessToken)
        this.setState({
          accessToken: response.accessToken
        })
      }).catch(error => {
        // fallback to interaction when silent call fails
        return msal.acquireTokenPopup(loginRequest)
          .then(response => {
            localStorage.setItem('accessToken', response.accessToken)
            this.setState({
              accessToken: response.accessToken
            })
          }).catch(error => {
            console.log(error);
          });
      });
  }

  signOut() {
    localStorage.removeItem('idToken')
    localStorage.removeItem('accessToken')
    msal.logout().then(() => {
      this.setState({
        idToken: '',
        accessToken: '',
        name: '',
        username: '',
        userPrincipalName: '',
        firstName: '',
        lastName: '',
        businessPhones: '',
        jobTitle: '',
        mail: '',
        mobilePhone: '',
        officeLocation: '',
        preferredLanguage: '',
        id: ''
      })
    })
  }

  async getProfileData() {
    let token = localStorage.getItem('accessToken')
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);

    const options = {
      method: "GET",
      headers: headers
    };

    fetch(graphConfig.graphMeEndpoint, options)
      .then(response => response.json())
      .then(data => {
        this.setState({
          name: data.displayName,
          username: data.userName,
          userPrincipalName: data.userPrincipalName,
          firstName: data.givenName,
          lastName: data.surname,
          businessPhones: data.businessPhones,
          jobTitle: data.jobTitle,
          mail: data.mail,
          mobilePhone: data.mobilePhone,
          officeLocation: data.officeLocation,
          preferredLanguage: data.preferredLanguage,
          id: data.id
        })
      })
      .catch(error => console.log(error))
  }

  render() {
    return (
      <center style={{ margin: '10vh' }}>
        <h2>Azure Active Directory Simulation with React</h2>
        <button disabled={msal.getAccount() ? true : false} onClick={() => this.signIn()}>Sign In</button>
        <button disabled={!localStorage.getItem('idToken')} onClick={() => this.getAccessToken()}>Get Access Token</button>
        <button disabled={!localStorage.getItem('accessToken')} onClick={() => this.getProfileData()}>Get Profile Data</button>
        <button disabled={!msal.getAccount()} onClick={() => this.signOut()}>Sign Out</button>
        <div>
          <span>ID Token:</span>
          <span className="token-field">
            :{this.state.idToken}
          </span>
        </div>
        <div>
          <span>Access Token:</span>
          <span className="token-field">
            :{this.state.accessToken}
          </span>
        </div>
        <table border="1" style={{ textAlign: 'left' }}>
          <tbody>
            <tr>
              <th>Name</th>
              <td>{this.state.name}</td>
            </tr>
            <tr>
              <th>Username</th>
              <td>{this.state.username}</td>
            </tr>
            <tr>
              <th>User Principal Name</th>
              <td>{this.state.userPrincipalName}</td>
            </tr>
            <tr>
              <th>First Name</th>
              <td>{this.state.firstName}</td>
            </tr>
            <tr>
              <th>Last Name</th>
              <td>{this.state.lastName}</td>
            </tr>
            <tr>
              <th>Business Phones</th>
              <td>{this.state.businessPhones}</td>
            </tr>
            <tr>
              <th>Job Title</th>
              <td>{this.state.jobTitle}</td>
            </tr>
            <tr>
              <th>Mail</th>
              <td>{this.state.mail}</td>
            </tr>
            <tr>
              <th>Mobile Phone</th>
              <td>{this.state.mobilePhone}</td>
            </tr>
            <tr>
              <th>Office Location</th>
              <td>{this.state.officeLocation}</td>
            </tr>
            <tr>
              <th>Preferred Language</th>
              <td>{this.state.preferredLanguage}</td>
            </tr>
            <tr>
              <th>ID</th>
              <td>{this.state.id}</td>
            </tr>
          </tbody>
        </table>
        <a href="https://www.npmjs.com/package/msal" target="_">NPM Package MSAL</a> <br />
      </center>
    );
  }
}

export default App;
