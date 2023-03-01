---8<--- "heading.md"

# Working with Teams Toolkit

## Overview

This is the extended lab where you will see how to make use of Teams Toolkit

---8<--- "are-you-on-the-right-path.md"

In this lab, once you finish setting up the Northwind Orders application, which can be found in the [A01-begin-app](https://github.com/microsoft/app-camp/blob/main/src/create-core-app/aad/A01-begin-app) folder, you will extend it further using Teams Toolkit.
Here, you will explore Teams Toolkit and create a Teams App using the Northwind Orders application as a Server component

* [A01-begin-app: Setting up the application with Azure AD](./A01-begin-app.md)
* [A02-after-teams-sso: Creating a Teams app with Azure ADO SSO](./A02-after-teams-sso.md)
* Working with Teams Toolkit (📍You are here)

In this lab you will learn to:

- Create a Tab App using Teams Toolkit
- Connect the Teams Toolkit with the Microsoft 365 Developer Sandbox
- Add new pages like Orders
- Fetch Orders from the Northwind Application
- Debug the Tab App on Teams

??? note "Table of Contents (open to display ►)"
    - [Overview](#overview)
    - [Features](#features)
    - [Exercise 1: Create new Tab App using Teams Toolkit](#exercise-1-create-new-tab-app-using-teams-toolkit)
      - [Step 1: Install Teams Toolkit Extension on VS Code](#step-1-install-teams-toolkit-extension-on-vs-code)
      - [Step 2: Create a new app using teams Toolkit](#step-2-create-a-new-app-using-teams-toolkit)
      - [Step 3: Connect Teams Toolkit with an existing Microsoft 365 Developer Account](#step-3-connect-teams-toolkit-with-an-existing-microsoft-365-developer-account)
      - [Step 4: Debug and see how the app runs](#step-4-debug-and-see-how-the-app-runs)
    - [Exercise 2: Modify App for Northwind Orders Application](#exercise-2-modify-app-for-northwind-orders-application)
      - [Step 1: Modify App.tsx](#step-1-modify-apptsx)
      - [Step 2: Add Dashboard.tsx](#step-2-add-dashboardtsx)
      - [Step 3: Add Dashboard.css](#step-3-add-dashboardcss)
      - [Step 4: Initializa Models](#step-4-initializa-models)
      - [Step 5: Create OrderGrids](#step-5-create-ordergrids)
      - [Step 6: Create Counters](#step-6-create-counters)
      - [Step 7: Connect Dashboard to tab](#step-7-connect-dashboard-to-tab)
      - [Step 8: Create API Endpoints](#step-8-create-api-endpoints)
    - [Exercise 3: Modify Northwind Orders Application](#exercise-3-modify-northwind-orders-application)
      - [Step 1: Enable CORS in the Application](#step-1-enable-cors-in-the-application)
      - [Step 2: Log AccessToken to Console](#step-2-log-accesstoken-to-console)


## Features

- Understand the offerings from Teams Toolkit
- Connect to Microsoft 365 Developer Sandbox
- Connect to Azure (optional)
- View Samples
- Create new App (even from samples)
- Extend this App to read data from Northwind Orders application
- Add pages to display data in a grid
- Add cards to display counter data on top


## Exercise 1: Create new Tab App using Teams Toolkit

You can complete these labs on a Windows, Mac, or Linux machine, but you do need the ability to install the prerequisites. If you are not permitted to install applications on your computer, you'll need to find another machine (or virtual machine) to use throughout the workshop.

### Step 1: Install Teams Toolkit Extension on VS Code

- Goto extensions and search for Teams Toolkit
- Install the extension from Microsoft

### Step 2: Create a new app using teams Toolkit

- Select Teams Tolkit extension from the Left side rail
- In the Teams Toolkit window, Create a new Teams App
- Select SSO-enabled Tab
- Once created, open the folder in VS Code 

### Step 3: Connect Teams Toolkit with an existing Microsoft 365 Developer Account

To debug the App in Teams, it is necessary to use the Microsoft 365 Developer Accounts. So let's connect them to Teams Toolkit extension

- In the Teams Toolkit extension, under Accounts section on top, select Sign in to Microsoft 365
- It will open a browser window for you to sign into the Microsoft 365 Developer Account
- Select any user, for example, enter the credentials for the user Adele Vance
- Once done, return back to VS Code and now you should be able to see the email id of the account you connected in place of Sign in to Microsoft 365 with a green check below stating Sideloading enabled
- This means you have successfully connected the Developer account

### Step 4: Debug and see how the app runs

To debug the App, simply click on the Debug button of Local environment under the Environments section.
This should launch the App in Teams

## Exercise 2: Modify App for Northwind Orders Application

You will modify the Teams Application code to enable connecting with Northwind Orders application and display the data

### Step 1: Modify App.tsx

- Locate App.tsx under the components directory
- Add `import React from 'React';` at the top of the page
- Add `export const ApiContext = React.createContext("");` before the constructor `export default function App()`
- Set the React Context for API URL pointing to the Northwind Orders application in the `let apiUrl = ""`
```
const getApiUrl = () => {
  let apiUrl = "https://teamsappcamp.loophole.site";
  return apiUrl != null ? apiUrl : "";
};

const [apiUrl, setApiUrl] = React.useState("");

React.useEffect(() => {
  const api = getApiUrl();
  setApiUrl(api);
}, []);
```
- Enclose the `<Router>` tag with `<ApiContext.Provider value={apiUrl}>`
- It should look like this
```
<TeamsFxContext.Provider
  value={{ theme, themeString, teamsUserCredential }}
>
  <Provider
    theme={theme || teamsTheme}
    styles={{ backgroundColor: "#eeeeee" }}
  >
    <ApiContext.Provider value={apiUrl}>
      <Router>
        <Route exact path="/">
          <Redirect to="/tab" />
        </Route>
        {loading ? (
          <Loader style={{ margin: 100 }} />
        ) : (
          <>
            <Route exact path="/privacy" component={Privacy} />
            <Route exact path="/termsofuse" component={TermsOfUse} />
            <Route exact path="/tab" component={Tab} />
            <Route exact path="/config" component={TabConfig} />
          </>
        )}
      </Router>
    </ApiContext.Provider>
  </Provider>
</TeamsFxContext.Provider>
```

### Step 2: Add Dashboard.tsx

- Add a new directory called pages in the components
- Add a new file Dashboard.tsx unders pages
- Paste the following code

```
import { useState, useContext, useEffect } from "react";
import { Flex,Image,Button,Label,Input} from "@fluentui/react-northstar";
import "./Dashboard.css";
import { useData } from "@microsoft/teamsfx-react";
import { ApiContext } from "../App";
import { Counter } from "../counters/Counter";
import { ApiEndpoints } from "../../lib/apiEndpoints";
import Employee from "../../models/employee";
import { TeamsFxContext } from "../Context";
import { OrderGrid } from "../ordergrids/OrderGrid";

export function Dashboard(props: {
  showFunction?: boolean;
  environment?: string;
}) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? "" : data!.displayName;

  const apiUrl = useContext(ApiContext);

  const [employee, setEmployee] = useState<Employee>();
  const [employeeId, setEmployeeId] = useState("");
  const [token, setToken] = useState("");
  const [employeeFound, setEmployeeFound] = useState(false);

  const findEmployee = async () => {
    if (employeeId !== "") {
      let url =
        apiUrl + ApiEndpoints.GET_EMPLOYEE + "?employeeId=" + employeeId;
      try {
        const response = await fetch(url, {
          method: "GET",
          headers: {
            "content-type": "application/json",
            Authorization: "Bearer " + token,
          },
        });
        if (response.status === 200) {
          const json = await response.json();
          console.log(json);
          setEmployee(json);
          setEmployeeFound(true);
        } else {
          setEmployeeFound(false);
          console.log("not found");
        }
      } catch (err) {
        console.log(err);
      }
    }
  };

  const saveToken = async (token: string) => {
    console.log(token);
    await setToken(token);
  };

  return (
    <div className="welcome page">
      <div className="narrow">
        <Image src="hello.png" />
        <h1 className="center">
          Congratulations{userName ? ", " + userName : ""}!
        </h1>
        <p className="center">
          Your app is running in your -{friendlyEnvironmentName}
        </p>
        <p className="center"> API endpoint: {apiUrl}</p>
        <p className="center">
          {" "}
          Your Access Token:
          <Input value={token}
            onChange={(e, v) => {
              saveToken(v?.value ? v.value : "");
            }}
            placeholder="Please enter your token"
          />
        </p>
        <Flex gap="gap.small" hAlign="center" vAlign="center">
          <Label content="Fill the required details." />
          <p className="center">
            {" "}
            Your Employee ID:
            <Input value={employeeId}
              onChange={(e, v) => {
                setEmployeeId(v?.value ? v.value : "");
              }}
              placeholder="Employee ID"
            />
          </p>
          <p className="center">
            Your mail address is:{" "}
            {employeeId !== "" && employee !== undefined
              ? employee?.mail
              : "Unauthorized"}{" "}
          </p>
          <Button primary content="Validate AAD" disabled={employeeFound} onClick={findEmployee} />
        </Flex>

        <div className="sections">
          {employeeFound && (
            <>
              <Flex gap="gap.small">
                <Counter title="Order(s)" type="order" orders={employee?.orders} />
                <Counter title="Customer(s)" type="customer" orders={employee?.orders} />
              </Flex>
              <OrderGrid employee={employee} />
            </>
          )}
        </div>
      </div>
    </div>
  );
}
```

### Step 3: Add Dashboard.css

- Add a new file Dashboard.css unders pages
- Paste the following code

```
.narrow {
  max-width: 900px;
  margin: 0 auto;
}

.page-padding {
  padding: 4rem;
}

.welcome.page > .narrow > img {
  margin: 0 auto;
  display: block;
  width: 200px;
}

.welcome.page > .narrow > ul {
  width: 80%;
  justify-content: space-between;
  margin: 4rem auto;
  border-bottom: none;
}

.welcome.page > .narrow > ul > li {
  background-color: inherit;
  margin: auto;
}

.welcome.page > .narrow > ul > li > a {
  font-size: 14px;
  min-height: 32px;
  border-bottom-color: rgb(98, 100, 167);
}

.center {
  text-align: center;
}

.sections > * {
  margin: 4rem auto;
}

pre,
div.error {
  background-color: #e5e5e5;
  padding: 1rem;
  box-shadow: 0px 1px 2px rgba(0, 0, 0, 0.2);
  border-radius: 3px;
  margin: 1rem 0;
  max-height: 200px;
  overflow-x: scroll;
  overflow-y: scroll;
}

pre.fixed,
div.error.fixed {
  height: 200px;
}

code {
  background-color: #e5e5e5;
  display: inline-block;
  padding: 0px 6px;
  border-radius: 3px;
  box-shadow: 0px 1px 2px rgba(0, 0, 0, 0.2);
}

.dark pre {
  background-color: #1b1b1b;
}

.dark code {
  background-color: #1b1b1b;
}

.dark div.error {
  background-color: #1b1b1b;
}

.error {
  color: red;
}

```

### Step 4: Initializa Models

- Create a new directory **models** under src directory
- Add a new file **order.ts** with the following code
```
class Order {
    customerContact: string
    customerId: string
    customerName: string
    customerPhone: string
    orderDate: string
    orderId: number
    shipAddress: string
    shipName: string
    shipRegion: string

    constructor(customerContact: string, customerId: string, customerName: string, customerPhone: string, orderDate: string, orderId: number, shipAddress: string, shipName: string, shipRegion: string) {
        this.customerContact = customerContact
        this.customerId = customerId
        this.customerName = customerName
        this.customerPhone = customerPhone
        this.orderDate = orderDate
        this.orderId = orderId
        this.shipAddress = shipAddress
        this.shipName = shipName
        this.shipRegion = shipRegion
    }
}

export default Order;
```
- Add a new file **employee.ts** with the following code
```
import Order from "./order"

class Employee {
    id?: number
    displayName?: string
    mail?: string
    jobTitle?: string
    city?: string
    photo?: string
    orders?: Order[]

    constructor(id: number, displayName: string, mail: string, jobTitle: string, city: string, photo: string, orders: Order[]) {
        this.id = id;
        this.displayName = displayName;
        this.mail = mail;
        this.jobTitle = jobTitle;
        this.city = city;
        this.photo = photo;
        this.orders = orders;
    }
}

export default Employee;
```
- Add a new file **product.ts** with the following code
```
import Order from "./order"

class Product {
    productId?: number
    productName?: string
    quantityPerUnit?: string
    supplierCountry?: string
    supplierName?: string
    unitPrice?: number
    orders?: Order[]
    categoryId: number
    reorderLevel: number
    categoryName?: string
    discontinued: boolean
    unitsInStock: number
    unitsOnOrder: number

    constructor(productId: number, productName: string, quantityPerUnit: string, supplierCountry: string, supplierName: string, unitPrice: number, orders: Order[], categoryId: number, categoryName: string, discontinued:boolean, reorderLevel: number, unitsInStock: number, unitsOnOrder: number) {
        this.productId = productId;
        this.productName = productName;
        this.quantityPerUnit = quantityPerUnit;
        this.supplierCountry = supplierCountry;
        this.supplierName = supplierName;
        this.unitPrice = unitPrice;
        this.orders = orders;
        this.categoryId = categoryId
        this.categoryName = categoryName
        this.discontinued = discontinued
        this.reorderLevel = reorderLevel
        this.unitsInStock = unitsInStock
        this.unitsOnOrder = unitsOnOrder
    }
}

export default Product;
```

### Step 5: Create OrderGrids

- Add a new directory called **ordergrids** under components
- Add a new file **OrderGrid.tsx** with the following code
```
import {
  ShorthandCollection,
  Table,
  TableRowProps,
} from "@fluentui/react-northstar";
import "./OrderGrid.css";
import Employee from "../../models/employee";
import Order from "../../models/order";

export function OrderGrid(props: { employee?: Employee }) {
  const header = {
    items: [
      {
        content: "Id",
        key: "Id",
        className: "_caseGrid_id _caseGrid_header",
      },
      {
        content: "Date",
        key: "Date",
        className: "_caseGrid_header",
      },
      {
        content: "Ship To",
        key: "ShipTo",
        className: "_caseGrid_header",
      },
      {
        content: "Address",
        key: "Address",
        className: "_caseGrid_header",
      },
    ],
  };

  let rows:
    | ShorthandCollection<TableRowProps, Record<string, {}>>
    | {
        key: number | undefined;
        items: (
          | {
              content: number | undefined;
              key: string;
              truncateContent: boolean;
              className: string;
            }
          | {
              content: string | undefined;
              key: string;
              truncateContent: boolean;
              className: string;
            }
          | {
              content: JSX.Element | undefined;
              key: string;
              truncateContent: boolean;
              className: string;
            }
        )[];
      }[]
    | undefined = [];

  if (props.employee) {
    console.log(props.employee);
    props.employee.orders?.forEach((item: Order) => {
      let row = {
        key: item.orderId,
        items: [
          {
            content: item.orderId,
            key: item.orderId + "-1",
            truncateContent: false,
            className: "_caseGrid_id",
          },
          {
            content: item.orderDate,
            key: item.orderDate + "-2",
            truncateContent: false,
            className: "",
          },
          {
            content: item.shipName,
            key: item.shipName + "-3",
            truncateContent: false,
            className: "",
          },
          {
            content: item.shipAddress,
            key: item.shipAddress + "-4",
            truncateContent: false,
            className: "",
          },
        ],
      };
      rows?.push(row);
    });
  }

  return (
    <div>
      <Table compact header={header} rows={rows} aria-label="Table" />
    </div>
  );
}
```
- Add a new file **OrderGrid.css** with the following code
```
._caseGrid_id {
  width: 50px !important;
  max-width: 50px !important;
}

._caseGrid_header {
  font-weight: bold;
}
```

### Step 6: Create Counters

- Add a new directory called **counters** under components
- Add a new file **Counter.tsx** with the following code
```
import { CardHeader, CardBody, Card, Flex, Text } from "@fluentui/react-northstar";
import { useEffect, useState } from "react";
import Order from "../../models/order";

export function Counter(props: {
  title?: string;
  type?: string;
  orders?: Order[];
}) {
  const [count, setCount] = useState(0);

  useEffect(() => {
    let tOrders = props.orders !== undefined ? props.orders : [];
    switch (props.type) {
      case "order":
        setCount(tOrders.length);
        break;
      case "customer":
        var unique: String[] = [];
        tOrders.forEach((order, index) => {
          if (unique.indexOf(order.customerId) === -1) {
            unique.push(order.customerId);
          }
        });
        setCount(unique.length);
        break;
    }
  }, [props]);

  return (
    <div>
      <Flex gap="gap.small">
        <Card
          aria-roledescription="card avatar"
          elevated
          inverted
          styles={{ height: "100px", width: "180px" }}
        >
          <Flex gap="gap.small" column fill vAlign="stretch" space="between">
            <CardHeader>
              <Text content={props.title} weight="bold" size="large" align="center"/>
            </CardHeader>
            <CardBody>
              <Text content={count} weight="bold" size="large" align="center" />
            </CardBody>
          </Flex>
        </Card>
      </Flex>
    </div>
  );
}
```

### Step 7: Connect Dashboard to tab

- Open Tab.tsx under components
- Add the import in the top `import { Dashboard } from "./pages/Dashboard";`
- Replace `<Welcome showFunction={showFunction} />` with `<Dashboard />`

### Step 8: Create API Endpoints

- Create a new directory **lib** under src
- Add a new file **apiEndpoints.tsx** with the following code
```
export class ApiEndpoints {
    static readonly GET_EMPLOYEE = "/api/employee";
    static readonly GET_ORDER = "/api/order";
    static readonly VALIDATE_LOGIN = "/api/validateAadLogin";
}
```
## Exercise 3: Modify Northwind Orders Application

To use this App as a Server Application for Teams Tab App, we will need to enable CORS and log the access token.
Once the modification is done, we will be able to use it directly with the Teams App.
_This is done only to ensure continuity._


### Step 1: Enable CORS in the Application
Since the Northwind Orders application, runs on it's own, it doesn't need to enable CORS.
However, in the Teams Application, you will make the API calls using **fetch** which requres the CORS Access policy to be implemented.

- Open the Northwind Orders application in a separate VS Code
- Open the server.js in server directoy
- Add the following code snippet after `const app = express();`
```
var allowedOrigins = ['https://localhost:53000', 'https://teamsappcamp.loophole.site', 'https://2084-101-0-63-152.in.ngrok.io']; // Include your Teams App URL
app.use(cors({
  origin: function (origin, callback) {
    // allow requests with no origin 
    // (like mobile apps or curl requests)
    if (!origin) return callback(null, true);
    if (allowedOrigins.indexOf(origin) === -1) {
      var msg = 'The CORS policy for this site does not ' +
        'allow access from the specified Origin.';
      return callback(new Error(msg), false);
    }
    return callback(null, true);
  }
}));
```

### Step 2: Log AccessToken to Console

- Open the file **identityClient.js** under the path `/client/identity`
- Locate the method `getLoggedinEmployeeId2()`
- Add this `console.log('Token Used: ', accessToken);` below `console.log(`Signed into account with employee ID ${data.employeeId}`);`
- Your **if** block should look like this
```
if (response.ok) {
    const data = await response.json();
    if (data.employeeId) {
        console.log(`Signed into account with employee ID ${data.employeeId}`);
        console.log('Token Used: ', accessToken);
        return data.employeeId;
    }
}
```

