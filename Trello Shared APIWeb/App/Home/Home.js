/// <reference path="../App.js" />
// global app
(function() {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function(reason) {
		$(document).ready(function() {
			app.initialize();

			// Event Handlers
			$('#deauthorize-with-trello').click(deauthorizeWithTrello);
			$('#authorize-with-trello').click(authorizeWithTrello);
			$('#get-boards-from-trello').click(getBoardsFromTrello);
			$('#get-selected-text').click(getSelectedText);
			
			// Initialize
			dataService = trelloDataService();	//sampleDataService();

			// Databind
			ko.applyBindings(vm);



		});
	};

	// Initialize View Models
	var vm = {
		isAuthenticated: false,
		boards: ko.observableArray([]),
		getCards: getCards
	};

	// Functions
	function showErrorMessage(errorMsg) {
		app.showNotification(errorMsg.responseText);
	}

	function deauthorizeWithTrello() {
		dataService.deauthorize();
		app.showNotification('Successful deauthorize');
	}

	function authorizeWithTrello() {
		dataService.authorize(function() {
				app.showNotification('Successful authentication');
			}, function() {
				app.showNotification('Failed authentication');
			}
		);
	}

	function getBoardsFromTrello() {
		dataService.getMyBoards(function(boards) {
			vm.boards.removeAll();
			boards = boards.filter(function(i, n) {
				return i.closed === false;
			});
			for (var index in boards) {
				vm.boards.push({
					id: boards[index].id,
					name: boards[index].name
				});
			}
		}, showErrorMessage);
	}

	function getCards(board) {
		dataService.getCardsForBoard(board, function(cards) {

			// Save current list for later
			vm.cards = cards;

			// Write card list to document			
			if (Office.CoercionType.Matrix) {
				// Convert to list		
				var matrixRow = [];
				cards.forEach(function(card) {
					matrixRow.push(card.name)
				})
				
				// Add list to Matrix
				var matrix = [];
				matrix.push(matrixRow);
				Office.context.document.setSelectedDataAsync(matrix, {coercionType:Office.CoercionType.Matrix},			 
					function (asyncResult) {
						if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
							// It worked, do nothing
						}
						if (asyncResult.status === Office.AsyncResultStatus.Failed) {
							var textToAdd = '';
							for (var i in cards) {
								var card = cards[i];
								var cardData = card.name + '\n' + card.desc + '\n';
								textToAdd = textToAdd + cardData;
							}
							Office.context.document.setSelectedDataAsync(textToAdd);
						}
				});
			} else {
				var textToAdd = '';
				for (var i in cards) {
					var card = cards[i];
					var cardData = card.name + '\n' + card.desc + '\n';
					textToAdd = textToAdd + cardData;
				}
				Office.context.document.setSelectedDataAsync(textToAdd);
			}		

		}, showErrorMessage);
	}
	
	function getSelectedText() {
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { }, function(asyncResult) {
			 callbackCheck(asyncResult, function(value) {
				 app.showNotification('Selected Text', value);
			 });
		});
	}
	
	function callbackCheck(asyncResult, success) {
		if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
			if (success)
				success(asyncResult.value);
		}
		if (asyncResult.status === Office.AsyncResultStatus.Failed) {
			var error = asyncResult.error;
			app.showNotification(error.code + ':' + error.name, error.message);
		}
	}
	

	// Services
	var dataService = {};

	function sampleDataService() {
		var ds = {
			authorize: function(success, failure) {
				if (success !== undefined) success();
			},
			deauthorize: function(success, failure) {
				if (success !== undefined) success();
			},
			getMyBoards: function(success, failure) {
				var boards = [{
					"id": "1",
					"name": "My Sample Board",
					"closed": false
				}, {
					"id": "2",
					"name": "My Second Sample Board",
					"closed": false
				}];
				success(boards);
			},
			getCardsForBoard: function(board, success, failure) {
				var cards = [{
					"name": "Card 1",
					"desc": "Description 1"
				}, {
					"name": "Card 2",
					"desc": "Description 2"
				}];

				success(cards);
			}
		};
		return ds;
	}
	
	function trelloDataService() {
		var ds = {
			authorize: function(success, failure) {
				Trello.authorize({
					// type: 'popup',
					name: 'Trello Shared API',
                    persist: true,
                    success: success,
                    error: failure
                });			
			},
			deauthorize: function(success, failure) {
				Trello.deauthorize();
			},			
			getMyBoards: function(success, failure) {
				Trello.get('/member/me/boards', success, failure);
			},
			getCardsForBoard: function(board, success, failure) {
				Trello.get('/boards/' + board.id + '/cards', success, failure);
			},
		};
		return ds;
	}
	$("#file").change(storeFileAsBase64);
	let chosenFileBase64;

	async function storeFileAsBase64() {
		const reader = new FileReader();

		reader.onload = async (event) => {
			const startIndex = reader.result.toString().indexOf("base64,");
			const copyBase64 = reader.result.toString().substr(startIndex + 7);

			chosenFileBase64 = copyBase64;
		};

		const myFile = document.getElementById("file") as HTMLInputElement;
		reader.readAsDataURL(myFile.files[0]);
	}
	async function insertAllSlides() {
		await PowerPoint.run(async function (context) {
			context.presentation.insertSlidesFromBase64(chosenFileBase64);
			await context.sync();
		});
	}
	async function insertSlidesDestinationFormatting() {
		await PowerPoint.run(async function (context) {
			context.presentation
				.insertSlidesFromBase64(chosenFileBase64,
					{
						formatting: "UseDestinationTheme",
						targetSlideId: "267#"
					}
				);
			await context.sync();
		});
	}
})();