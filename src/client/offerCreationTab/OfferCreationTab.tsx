import * as React from "react";
import Axios from "axios";
import { Provider, Flex, Text, Button, Header, Loader } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import { OfferCreationForm } from "./OfferCreationForm";
import{ IOffer } from '../../model/IOffer';

/**
 * Implementation of the Offer Creation content page
 */
export const OfferCreationTab = () => {
  const [{ inTeams, theme, context }] = useTeams();
  const [idToken, setIdToken] = useState<string>();
  const [showSpinner, setShowSpinner] = useState<boolean>(false);

  const createOffer = (offer: IOffer) => {
    if (idToken) {
      setShowSpinner(true);
      const requestBody = {
        domain: context?.sharePointSite?.teamSiteDomain,
        offer: offer
      };
      Axios.post(`https://${process.env.PUBLIC_HOSTNAME}/api/createoffer`, requestBody, {
                  responseType: "json",
                  headers: {
                    Authorization: `Bearer ${idToken}`
                  }
      }).then(result => {
        console.log(result);        
      })
      .catch((error) => {
        console.log(error);
      })
      .finally(() => {
        setShowSpinner(false);
      });
    }
  };

  useEffect(() => {
    if (inTeams === true) {
      authentication.getAuthToken({
          resources: [`api://${process.env.PUBLIC_HOSTNAME}/${process.env.TAB_APP_ID}`],
          silent: false
      } as authentication.AuthTokenRequestParameters).then(token => {
          setIdToken(token);
          app.notifySuccess();
      }).catch(message => {
          app.notifyFailure({
              reason: app.FailedReason.AuthFailed,
              message
          });
      });
    } else {
      //  setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="This is your tab" />
                </Flex.Item>
                <Flex.Item>
                    <OfferCreationForm createOffer={createOffer} />
                </Flex.Item>
                <Flex.Item>
                    <div>{showSpinner && <Loader label="Creating document" labelPosition="above" />}</div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Markus Moeller" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
