import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import MsgSearchCommandMessageExtension from "../msgSearchCommandMessageExtension/MsgSearchCommandMessageExtension";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler } from "botbuilder";



// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for MsgSearchCommand Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class MsgSearchCommandBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    /** Local property for MsgSearchCommandMessageExtension */
    @MessageExtensionDeclaration("msgSearchCommandMessageExtension")
    private _msgSearchCommandMessageExtension: MsgSearchCommandMessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        // Message extension MsgSearchCommandMessageExtension
        this._msgSearchCommandMessageExtension = new MsgSearchCommandMessageExtension();


        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);


    }


}
