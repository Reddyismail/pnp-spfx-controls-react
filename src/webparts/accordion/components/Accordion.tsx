import * as React from 'react';
import styles from './Accordion.module.scss';
import type { IAccordionProps } from './IAccordionProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import { Accordion } from '@pnp/spfx-controls-react';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from "@pnp/spfx-controls-react/lib/AccessibleAccordion";

export default class Accordions extends React.Component<IAccordionProps> {
  public render(): React.ReactElement<IAccordionProps> {
    const {
      isDarkTheme,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.accordion} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>
            <Accordion 
            // allowMultipleExpanded={true}
            allowZeroExpanded={true}
            preExpanded={["Hallow"]}
            >
              <AccordionItem>
                <AccordionItemHeading>
                  <AccordionItemButton>
                    What harsh truths do you prefer to ignore?
                  </AccordionItemButton>
                </AccordionItemHeading>
                <AccordionItemPanel>
                  <p>
                    Exercitation in fugiat est ut ad ea cupidatat ut in
                    cupidatat occaecat ut occaecat consequat est minim minim
                    esse tempor laborum consequat esse adipisicing eu
                    reprehenderit enim.
                  </p>
                </AccordionItemPanel>
              </AccordionItem>
              <AccordionItem>
                <AccordionItemHeading>
                  <AccordionItemButton>
                    Is free will real or just an illusion?
                  </AccordionItemButton>
                </AccordionItemHeading>
                <AccordionItemPanel>
                  <p>
                    In ad velit in ex nostrud dolore cupidatat consectetur
                    ea in ut nostrud velit in irure cillum tempor laboris
                    sed adipisicing eu esse duis nulla non.
                  </p>
                </AccordionItemPanel>
              </AccordionItem>
            </Accordion>

          </div>
        </div>
      </section>
    );
  }
}
