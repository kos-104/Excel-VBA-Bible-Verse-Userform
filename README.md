The Excel VBA Bible Userform is an internet tool that retrieves Bible verses and displays them in Excel. It employs the Model-View-ViewModel (MVVM) architecture to enhance maintainability and scalability. This structure is modeled after the "RubberduckSwagShop MVVM-Lite" project. The design pattern separates the user interface from business logic, allowing flexible data binding and command execution.

Key features include the ability to resize the form and all controls, as well as adjust the font size of the downloaded text verse(s) - facilitating a responsive user experience across different screen sizes. The userform offers intuitive navigation with tooltips and spin buttons, improving user interaction. 

Users can download multiple verses from the same chapter at once, enhancing efficiency. Listboxes in the userform allow mouse scrolling for added convenience. Command buttons also change colors when the mouse hovers over them, helping users pinpoint the button targeted. Event handling is thorough, with custom commands for adding, clearing, and confirming line items which are supported by error handling.

Special acknowledgment goes to David Hager for his extensive groundwork on the dynamic named ranges/formulas and pivot tables that underpin the form's framework (https://dhexcel1.wordpress.com/2017/07/03/lookup-a-bible-verse-using-excel-wo-vba-by-david-hager/). An additional thanks to Mathieu Guindon and the Rubberduck team for creating the “SwagShop” MVVM-Lite program utilized to build this project (https://rubberduckvba.blog/2023/04/11/lightweight-mvvm-in-vba/).



