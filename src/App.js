import React, { useState } from 'react';

const App = props => {

  let [greeting, setGreeting] = useState("Yoooo")


    return (
      <div className="App">
        <p>{greeting}</p>
      </div>
    );
}

export default App;
