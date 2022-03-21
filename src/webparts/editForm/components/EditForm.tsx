import * as React from 'react';
import { IEditFormProps } from './IEditFormProps';
import { IEditFormStates } from './IEditFormStates';
import { escape } from '@microsoft/sp-lodash-subset';
import { Container, Typography, FormControl, TextField, Select, MenuItem, Button } from '@material-ui/core';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export default class EditForm extends React.Component<IEditFormProps, IEditFormStates, any> {
  constructor(props) {
    super(props);

    this.state = {
      taskID: '',
      task: '',
      description: '',
      priority: '',
      dueDate: null,
      taskExecutor: [],
      emailTaskExecutor: '',
      urlId: null
    };
  }

  componentDidMount = () => {
    this.getPageId();
    console.log('33');
  }

  componentDidUpdate(prevProps: Readonly<IEditFormProps>, prevState: Readonly<any>, snapshot?: any): void {
    console.log('Updated', this.state);
  }

  getPageId = () => {
    let url : any = new URL(window.location.href);
    let formId = url.searchParams.get("FormID");
    console.log('formId', formId);
    this.setState({ urlId: formId }, () => {
      this.fetchData();
      console.log('urlId', this.state.urlId);
    });
  }
  
  fetchData = () => {
    let web = Web(this.props.webURL);
    web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.getById(this.state.urlId)
      .get()
      .then(result => {
        console.log('result', result);
        this.setState({
          taskID: result.ID,
          task: result.Title,
          description: result.Description,
          priority: result.Priority,
          dueDate: result.Due_x0020_date.slice(0,10),
          taskExecutor: [result.Email_x0020_Task_x0020_Executor]
        });
      }).catch(Err => {
        console.error(Err);
      });
  }

  handleChange = (e: {target: {name: any; value: any; }; }) => {
    // console.log(e.target.value);
    const newState = { [e.target.name]: e.target.value } as Pick<IEditFormStates, keyof IEditFormStates>;
    this.setState(newState);
  }

  handleSelectChange = (e) => {
    // console.log('e.target',e.target.value);
    this.setState({ priority: e.target.value });
  }

  getPeoplePickerItems = (items: any[]) => {
    console.log('Items:', items);
    this.setState({taskExecutor: items, emailTaskExecutor: items[0].secondaryText});
    // this.setState({taskExecutor: items});
  }

  updateItem = () => {
    let web = Web(this.props.webURL);
    web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.getById(this.state.urlId).update({
      Title: this.state.task,
      Description: this.state.description,
      Priority: this.state.priority,
      Due_x0020_date: this.state.dueDate,
      Task_x0020_ExecutorId: this.state?.taskExecutor[0]?.id,
      Email_x0020_Task_x0020_Executor: this.state.emailTaskExecutor
    }).then(result => {
      alert("Item Updated Successfully");
      location.href = 'https://projects1.sharepoint.com/sites/Development/Alona/Lists/NewTasks/AllItems.aspx?viewpath=%2Fsites%2FDevelopment%2FAlona%2FLists%2FNewTasks%2FAllItems%2Easpx';
    })
    console.log('state', this.state);
  }


  public render(): React.ReactElement<IEditFormProps> {
    return (
      <Container maxWidth="sm">
        <Typography variant="h6" style={{ textAlign: 'center', marginTop: '10px', marginBottom: '10px' }}>Edit Task</Typography>
        <FormControl style={{ marginTop: '20px' }}>
          <TextField label="Task" name="task" value={this.state.task} variant="outlined" onChange={this.handleChange} style={{ marginTop: '13px', width: '500px' }}/>
          <TextField label="Task Description" name="description" value={this.state.description} variant="outlined" onChange={this.handleChange} multiline rows={3} style={{ marginTop: '13px', width: '500px' }}/>
          <Select
            label="Priority"
            name="priority"
            value={this.state.priority}
            variant="outlined" 
            onChange={(e) => {this.handleSelectChange(e)}}
            style={{ marginTop: '13px', width: '500px' }}
          >
            <MenuItem value="High">High</MenuItem>
            <MenuItem value="Medium">Medium</MenuItem>
            <MenuItem value="Low">Low</MenuItem>
          </Select>
          <TextField
            id="date"
            variant="outlined"
            label="Due Date"
            type="date"
            name="dueDate"
            value={this.state.dueDate}
            InputLabelProps={{
              shrink: true,
            }}
            style={{ marginTop: '13px', width: '500px' }} 
            onChange={this.handleChange}
          />
          <PeoplePicker
            context={this.props.context as any}
            titleText="Task Executor"
            groupName={''}
            personSelectionLimit={1}
            required={false}
            showHiddenInUI={false}
            defaultSelectedUsers={this.state.taskExecutor}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            ensureUser={true}
            onChange={this.getPeoplePickerItems}
          />
        </FormControl>
        <div style={{ marginTop: '20px' }}>
          <Button style={{ width: '83px', marginRight: '5px'}} variant="outlined" onClick={this.updateItem} color="primary">Update</Button>
        </div>
      </Container>
    );
  }
}
