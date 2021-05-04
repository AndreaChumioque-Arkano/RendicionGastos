// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './components/App';
import { Provider, teamsTheme } from '@fluentui/react-northstar' //https://fluentsite.z22.web.core.windows.net/quick-start

ReactDOM.render(
    <Provider themeName={teamsTheme} lang="en-US">
        <App />
    </Provider>, document.getElementById('root')
);
