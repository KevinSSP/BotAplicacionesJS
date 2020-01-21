// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');

class MyBot extends ActivityHandler {
    constructor(configuration, qnaOptions) {
        super();
        if (!configuration) throw new Error('[QnaMakerBot]: Falta la ocnfiguracion del Servicio de Preguntas y Respuestas');
        // now create a qnaMaker connector.
        this.qnaMaker = new QnAMaker(configuration, qnaOptions);
        //
        this.onMembersAdded(async (context, next) => {
            await this.sendWelcomeMessage(context);
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            const text = context.activity.text;

            // Create an array with the valid color options.
            const validColors = ['APLICACIONES FINANCIERAS', 'COMPRAS Y COMERCIO EXTERIOR', 'PEDIDOS Y PRECIOS', 'INVENTARIO, WMS Y DESPACHOS', 'FACTURACION, CARTERA Y CRM', 'PLANEACION Y COSTOS', 'PRODUCCION'];

            if (validColors.includes(text)) {
                await context.sendActivity('Escribe tu pregunta, te brindaremos una posible solucion.');
            } else {
                // await context.sendActivity('Seleccione una opcion valida del menu');
                // await this.sendSuggestedActions(context);
                await this.sendAnswerQnA(context);
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    /**
     * Send a welcome message along with suggested actions for the user to click.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendWelcomeMessage(turnContext) {
        const { activity } = turnContext;

        // Iterate over all new members added to the conversation.
        for (const idx in activity.membersAdded) {
            if (activity.membersAdded[idx].id !== activity.recipient.id) {
                const welcomeMessage = 'Bienvenido al Bot de soporte de aplicaciones de Carvajal Tecnologia y Servicios' +
                    '. Seleccione una opcion:';
                await turnContext.sendActivity(welcomeMessage);
                await this.sendSuggestedActions(turnContext);
            }
        }
    }

    /**
     * Send suggested actions to the user.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendSuggestedActions(turnContext) {
        var reply = MessageFactory.suggestedActions(['APLICACIONES FINANCIERAS', 'COMPRAS Y COMERCIO EXTERIOR', 'PEDIDOS Y PRECIOS', 'INVENTARIO, WMS Y DESPACHOS', 'FACTURACION, CARTERA Y CRM', 'PLANEACION Y COSTOS', 'PRODUCCION']);
        await turnContext.sendActivity(reply);
    }

    /**
     * QnA Maker Respuestas
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendAnswerQnA(turnContext) {
        // send user input to QnA Maker.
        const qnaResults = await this.qnaMaker.getAnswers(turnContext);

        // If an answer was received from QnA Maker, send the answer back to the user.
        if (qnaResults[0]) {
            await turnContext.sendActivity(qnaResults[0].answer);
        } else {
            // If no answers were returned from QnA Maker, reply with help.
            await turnContext.sendActivity('No hemos encontrado una respuesta a tu solicitud.');
        }
    }
}

module.exports.MyBot = MyBot;
