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
      task: '',
      description: '',
      priority: '',
      dueDate: null,
      taskExecutor: [],
      emailTaskExecutor: ''
    };
  }

  componentDidMount() {
    this.fetchData();
  }

  async fetchData() {
    let web = Web(this.props.webURL);
    const items: any[] = await web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.get();
    console.log('items', items);
  }

  public render(): React.ReactElement<IEditFormProps> {
    return (
      <Container maxWidth="sm">
        <Typography variant="h6" style={{ textAlign: 'center', marginTop: '10px', marginBottom: '10px' }}>New Task</Typography>
        <FormControl style={{ marginTop: '20px' }}>
          <TextField label="Task" name="task" value={this.state.task} variant="outlined" style={{ marginTop: '13px', width: '500px' }}/>
          <TextField label="Task Description" name="description" value={this.state.description} variant="outlined" multiline rows={3} style={{ marginTop: '13px', width: '500px' }}/>
          <Select
            label="Priority"
            name="priority"
            value={this.state.priority ? this.state.priority : 'Low'}
            
            variant="outlined" 
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
            // value={this.state.dueDate ? this.state.dueDate : (new Date().toJSON().slice(0,10))}
            value={this.state.dueDate ? this.state.dueDate : ''}
            InputLabelProps={{
              shrink: true,
            }}
            style={{ marginTop: '13px', width: '500px' }}
            
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
            
          />
        </FormControl>
        <div style={{ marginTop: '20px' }}>
          <Button style={{ width: '83px', marginRight: '5px'}} variant="outlined" color="primary">Update</Button>
          <Button variant="outlined" color="secondary" >Cancel</Button>
        </div>
      </Container>
    );
  }
}
