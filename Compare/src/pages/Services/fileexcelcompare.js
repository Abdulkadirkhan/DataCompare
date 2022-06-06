import {
    Box,
    FormControl,
    InputLabel,
    Select,
    MenuItem,
    Button,
    Paper,Alert , AlertTitle
  } from "@mui/material";
  import React, { useState } from "react";
  import axios from "axios";
  
  
  const Upload = () => {
    const [filetype1, setFiletype1] = useState("");
    const [filetype2, setFiletype2] = useState("");
    const [success, setSuccess] = useState(false);
    const [error, setError] = useState(false);
    const [fileName1, setFileName1] = useState('');
    const [fileName2, setFileName2] = useState('');
  
    const handleChange = (event) => {
      console.log('handle change', event.target.files[0].name)
      setFileName1(event.target.files[0].name);
      setFiletype1(event.target.value);
    };

    const handleChange1 = (event) => {
        console.log('handle change1', event.target.files[0].name)
        setFileName2(event.target.files[0].name);
        setFiletype2(event.target.value);
      };
  
    const handleSubmit = (event) => {
      event.preventDefault();
      const data = new FormData(event.currentTarget);
  
      // console.log('data', data.get('filetype'))
  
      const baseUrl = "";
  
      // fetch(baseUrl+'/upload').then((response) => {
      //   response.json().then((body) => {
      //     console.log('file response +++++++++++++++++++++++++++++++', body)
      //     this.setState({ imageURL: `http://localhost:8000/${body.file}` });
      //   });
      // });
  
      axios
        .post("http://localhost:5000/singleUpload", data, {
          method: "GET",
          responseType: "blob",
        })
        .then((response) => {
          const url = window.URL.createObjectURL(new Blob([response.data]));
          const link = document.createElement("a");
          link.href = url;
          link.setAttribute("download", "Result.xlsx");
          document.body.appendChild(link);
          link.click();
          setSuccess(true)
        })
        .catch((error) => {
          console.log(error);
          setError(true);
          setSuccess(false)
        });
    };
  
    return (
        
        
      <Box sx={{ marginTop: '11rem', marginLeft: '29rem' }}>
           <div>
            <h2 style={{marginBottom : '2rem'}}>
              One to One File Comparsion
            </h2>
            <h3>
              Please Upload files to compare!
            </h3>
           </div>
        <Box component="form" onSubmit={handleSubmit} noValidate sx={{ mt: 1 }}>
        <Box>
        <Button sx={{ mt: 3, mb: 2 }} variant="contained" component="label">
            Select first file
            <input type="file" hidden name="file1" onChange={handleChange}/>
        </Button> 
        <Button sx={{ mt: 3, mb: 2, marginLeft: '10rem' }} variant="contained" component="label">
            Select second file
            <input type="file" hidden name="file2" onChange={handleChange1}/>
        </Button>       
        </Box>
        <Box sx={{    marginLeft: '7rem'}}>
        <Button
          type="submit"
          disabled= {filetype1 & filetype2 == '' ? true : false}
          variant="contained"
          sx={{ mt: 4, mb: 2, ml: 2, marginLeft: '4rem' }}
          style={{
            maxWidth: "500px"
            
          }}
        >Start Compare              
          
        </Button>      
        </Box>  
  
          <br></br>
          {fileName1 != '' && 
          <Box sx={{width : '40%', marginTop: '1rem'}}>
          
          <Alert severity="info">You have successfully uploaded file :  {fileName1}</Alert>
        </Box>
          }
           {fileName2 != '' && 
          <Box sx={{width : '40%', marginTop: '1rem'}}>
          
          <Alert severity="info">You have successfully uploaded file :  {fileName2}</Alert>
        </Box>
          }
          
          <Box sx={{width : '40%', marginTop: '2rem'}}>
          {success &&
  
            <Alert severity="success">
              <AlertTitle>Success</AlertTitle>
              This is a success alert — <strong>check it out!</strong>
            </Alert>
  
          }
          {error &&
  
            <Alert severity="error">
              <AlertTitle>Success</AlertTitle>
              This is a error alert — <strong>check it out!</strong>
            </Alert>
  
          }
  </Box>
        </Box>
      </Box>
    );
  };
  
  export default Upload;
  