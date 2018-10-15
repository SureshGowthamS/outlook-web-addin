import React from 'react';
import {render} from 'react-dom';
import MonacoEditor from 'react-monaco-editor/lib/editor';
import * as desktopmessagereadmode from './desktopmessagereadmode.json';
import * as desktopmessagecomposemode from './desktopmessagecomposemode.json';
import * as desktopcalendarreadmode from './desktopcalendarreadmode.json';
import * as desktopcalendarcomposemode from './desktopcalendarcomposemode.json';
import './index.css';
const Office = window.Office;

class App extends React.Component {
    constructor(props) {
        super(props);
        Office.initialize = () =>this.initData(); 
        this.state = {
            data : desktopmessagereadmode,
            api: "displayLanguage",
            code: "console.log(Office.context.displayLanguage);",
            component: "Office.context",
            language: "javascript"
        }
    }

    initData() {
        const readMode = (Office.context.mailbox.item.displayReplyAllForm !== undefined);
        var newData = this.state.data;
        if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
            newData = readMode ? desktopmessagereadmode : desktopmessagecomposemode;
        } else {
            newData = readMode ? desktopcalendarreadmode : desktopcalendarcomposemode;
        } 
        const newState = {
            data: newData,
            language: this.state.language,
            code: this.state.code,
            api: this.state.api,
            component: this.state.component
        };
        this.setState(newState);
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => this.handleOnItemChange());
    }

    handleOnAPIChange(newAPI) {
        const newCode = this.state.data["default"][this.state.component][newAPI];
        const newState = {
            data: this.state.data,
            language: "javascript",
            code: newCode,
            api: newAPI,
            component: this.state.component
        };
        this.setState(newState);
    }

    handleOnCodeChange(newCode) {
        const newState = {
            data: this.state.data,
            language: "javascript",
            code: newCode,
            api: this.state.api,
            component: this.state.component
        };
        this.setState(newState);
    }

    handleOnComponentChange(newComponent) {
        const newAPI = Object.keys(this.state.data["default"][newComponent])[0];
        const newCode = this.state.data["default"][newComponent][newAPI];
        const newState = {
            data: this.state.data,
            language: "javascript",
            code: newCode,
            api: newAPI,
            component: newComponent
        };
        this.setState(newState);
    }

    handleOnItemChange() {
        if (this.state.language === "json") {
            this.handleOnItemDataChange();
        }
    }

    handleOnItemDataChange() {
        var newCode = "";
        try {
            var element = document.createElement("pre");
            element.innerHTML = JSON.stringify(Office.context.mailbox.item, null , 4);
            newCode = element.innerHTML;
        } catch(err) {
            newCode = err.message;
        }
        const newState = {
            data: this.state.data,
            language: "json",
            code: newCode,
            api: this.state.api,
            component: this.state.component
        };
        this.setState(newState);
    }

    handleOnPersistentEvent() {
        console.log("Pin the add-in and switch mail items to view all properties");
        this.handleOnItemDataChange();
    }

    handleOnRunCode() {
        try {
            eval(this.state.code);
        }
        catch (err) {
            console.log(err.message);
        }
    }

    render() {
        return (
            <div className="app">
                <Title
                    api={this.state.api}
                    component={this.state.component}
                    apis={Object.keys(this.state.data["default"][this.state.component])}
                    components={Object.keys(this.state.data["default"])}
                    onAPIChange={(newAPI) => this.handleOnAPIChange(newAPI)}
                    onComponentChange={(newComponent) => this.handleOnComponentChange(newComponent)} />
                <Editor
                    language={this.state.language}
                    code={this.state.code}
                    onChange={(newCode) => this.handleOnCodeChange(newCode)} />
                <Console />
                <Actions
                    onRunCode={() => this.handleOnRunCode()}
                    onPersistent={() => this.handleOnPersistentEvent()}
                />
            </div>
        );
    }
}

class Title extends React.Component {

    handleOnAPIChange(newAPI) {
        if (this.props && this.props.onAPIChange) {
            this.props.onAPIChange(newAPI);
        }
    }

    handleOnComponentChange(newComponent) {
        if (this.props && this.props.onComponentChange) {
            this.props.onComponentChange(newComponent);
        }
    }

    render() {
        var componentOptions = [];
        if (this.props && this.props.components) {
            this.props.components.forEach((option, index) => {
                componentOptions.push(<option key={index} value={option}>{option}</option>);
            });
        }
        var apiOptions = [];
        if (this.props && this.props.apis) {
            this.props.apis.forEach((option, index) => {
                apiOptions.push(<option key={index} value={option}>{option}</option>);
            });
        }
        return (
            <div className="title">
                <select className="components" value={this.props.component} onChange={(event) => this.handleOnComponentChange(event.target.value)}>
                    {componentOptions}
                </select>
                <br />
                <select className="apis" value={this.props.api} onChange={(event) => this.handleOnAPIChange(event.target.value)}>
                    {apiOptions}
                </select>
            </div>
        );
    }
}

class Editor extends React.Component {
    constructor(props) {
        super(props);
        this.editor = null;
    }
    editorDidMount(editor) {
        editor.focus();
        this.editor = editor;
        window.addEventListener('resize', () => this.handleResize());
    }

    handleResize() {
        this.editor.layout();
    }

    handleOnChange(newValue) {
        if (this.props && this.props.onChange) {
            this.props.onChange(newValue);
        }
    }

    render() {
        const code = (this.props && this.props.code) ? this.props.code : '// type your code...';
        const options = {
            minimap: {
                enabled: false
            },
            lineNumbersMinChars: 2,
            autoIndent: true,
            formatOnPaste: true,
            formatOnType: true,
            folding: false
        };
        return (
            <div className="editor">
                <MonacoEditor
                    language={this.props.language}
                    theme="vs-light"
                    options={options}
                    value={code}
                    onChange={(newValue) => this.handleOnChange(newValue)}
                    editorDidMount={(editor, monaco) => this.editorDidMount(editor)}
                />
            </div>
        );
    }
}

class Actions extends React.Component {
    handleOnRunCode() {
        if (this.props && this.props.onRunCode) {
            this.props.onRunCode();
        }
    }

    handleOnPersistent() {
        if (this.props && this.props.onPersistent) {
            this.props.onPersistent();
        }
    }

    render() {
        return (
            <div className="actions">
                <button onClick={() => this.handleOnRunCode()}>Run code</button>
                <button onClick={() => this.handleOnPersistent()}>All Properties</button>
            </div>
        );
    }
}

class Console extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            visible: false,
            message: ""
        };
        console.log = (message) => this.handleOnMessage(message);
    }

    handleOnClose() {
        const newState = {
            visible: false,
            message: ""
        };
        this.setState(newState);
    }

    handleOnMessage(newMessage) {
        if (newMessage === undefined) {
            newMessage = "undefined";
        } else {
            newMessage = newMessage.toString();
        }
        const newState = {
            visible: true,
            message: newMessage
        };
        this.setState(newState);
    }

    render() {
        const element = this.state.visible ?
            (<div className="console">
                <div>{this.state.message}</div>
                <button className="closeButton" onClick={() => this.handleOnClose()}>X</button>
            </div>)
            : null;
        return (element);
    }
}
render(<App />, document.getElementById('app'));

