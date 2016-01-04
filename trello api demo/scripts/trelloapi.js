/* global $ */
var ConnectToTrello = function () {
    //console.log("Authenticating");
    Trello.authorize({
        type: 'popup',
        name: 'Outlook Trello Add-In',
        scope: { read: true, write: true, account: true },
        success: authenticationSuccess,
        error: authenticationError
    });
};

var GetTrelloLists = function () {
    if(console)console.log("Updating list of lists.");

    $('#trelloLists').empty();
    $('#trelloLists').append("<option value='false'>Select a List</option>");



    var boardId = $('#trelloBoards').val();
    if (boardId == "false") {
        return;
    }



    Trello.get('/boards/' + boardId + '/lists',
        function (success) {

            for (var i = success.length - 1; i >= 0; i--) {
                var list = success[i];
                var listOption = document.createElement("option");
                listOption.value = list.id;
                listOption.innerText = list.name;
                $('#trelloLists').append(listOption);
            };
            $("#TrelloListsPanel").show();
            //$("#trelloListsDropdown").Dropdown();
        },
        function (error) {
           if(console)console.log(error);
            

        }
    );
};

var GetTrelloCards = function () {
    $('#trelloCards').empty();

    var listId = $('#trelloLists').val();

    if (listId == "false") {
        return;
    }


    Trello.get('/lists/' + listId + '/cards',
        function (success) {


            for (var i = success.length - 1; i >= 0; i--) {
                var card = success[i];

                // var cardHTML = "<div class=\"panel panel-default trellocard\" onclick=\"InsertCard('" + card.id + "')\"><div class=\"panel-body\">"+card.name+"</div></div>";
                var cardHTML="<div class=\"ms-ListItem is-selectable\"><span class=\"ms-ListItem-primaryText\">" + card.name + "</span><span class=\"ms-ListItem-secondaryText\">card title</span> <span class=\"ms-ListItem-tertiaryText\">extra card information or description</span><div class=\"ms-ListItem-selectionTarget js-toggleSelection\" data-id=\""+ card.id +"\"></div></div>";
                
                $('#trelloCards').append(cardHTML);


            };
            $("#TrelloCardsPanel").show();
            InitSelectableListItems();
        },
        function (error) {

            console.log(error);

        }
    );
};

var InsertCard = function (cardId) {
    Trello.get('/cards/' + cardId,
			function (card) {
                alert(card.desc);

// 			    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync("RE: " + card.name);
// 			    Office.context.mailbox.item.body.setSelectedDataAsync("<div style=\"border-left-width: 2px;border-left-color: #0067A3;border-left-style: solid;padding-left: 10px;\"><h2>" + card.name + "</h2><div>" + card.desc + "</div></div>", { coercionType: Office.CoercionType.Html });
// 
// 			    app.showNotification(card.name, "Trello card added to message!");
// 


			    //$('#cardName').text(success.name + "(" + cardId + ")");
			    //$('#cardDescription').text(success.desc);
			    //$('#cardLink').attr("href", success.shortUrl);
			},
			function (error) {
			    // app.showNotification("Error getting card", error);

			}
		);
};

var authenticationSuccess = function () {
    
    $('#TrelloAuthenticationPanel').hide();
    
    // Load Boards
    Trello.get('/member/me/boards',
		function (data) {
		    for (var i = data.length - 1; i >= 0; i--) {
		        var board = data[i];
		        var boardOption = document.createElement("option");
		        boardOption.value = board.id;
		        boardOption.innerText = board.name;
                $('#trelloBoards').append(boardOption);
		    };
		    $("#TrelloBoardsPanel").show();
            //$("#trelloBoardsDropdown").Dropdown();
		},
		function (error) {
		    if(console)console.log(error);
		    //app.showNotification("Error loading boards", error);
		}
	);

}

var authenticationError = function (error) {
    console.log(error);
   

}


function InitSelectableListItems(){
    $('.js-toggleSelection').each(function(){
        $(this).on('click',function(){
             $(this).parents('.ms-ListItem').toggleClass('is-selected');
             
            //Insert into email when selected
            if( $(this).parents('.ms-ListItem').hasClass('is-selected')){
                //get the card Id
                var cardId=$(this).data('id');
                if(cardId != null)InsertCard(cardId);
            }
        
        });
    });
    
    
}
  