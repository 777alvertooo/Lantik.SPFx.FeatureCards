import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './FeatureCards.module.scss';
import type { IFeatureCardsProps } from './IFeatureCardsProps';
import type { IFeatureCardItem } from '../FeatureCardsWebPart';

export default class FeatureCards extends React.Component<IFeatureCardsProps, {}> {

  private _handleCardClick = (tarjeta: IFeatureCardItem): void => {
    if (tarjeta.enlace) {
      window.open(tarjeta.enlace, '_blank', 'noopener,noreferrer');
    }
  }

  public render(): React.ReactElement<IFeatureCardsProps> {
    const {
      tituloSeccion,
      tarjetas,
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${styles.featureCards} ${hasTeamsContext ? styles.teams : ''}`}>
        {tituloSeccion && (
          <h2 className={styles.sectionTitle}>{tituloSeccion}</h2>
        )}

        {!tarjetas || tarjetas.length === 0 ? (
          <div className={styles.emptyState}>
            Edita el Web Part y pulsa en <strong>Gestionar tarjetas</strong> en el panel derecho para añadir elementos.
          </div>
        ) : (
          <div className={styles.grid}>
            {tarjetas.map((tarjeta, index) => (
              <article
                key={index}
                className={`${styles.card} ${tarjeta.enlace ? styles.clickable : ''}`}
                onClick={() => this._handleCardClick(tarjeta)}
              >
                <div className={styles.iconBox}>
                  <Icon iconName={tarjeta.icono} />
                </div>
                <div className={styles.textContent}>
                  <h3 className={styles.cardTitle}>{tarjeta.texto}</h3>
                  <p className={styles.cardSubtitle}>{tarjeta.subtexto}</p>
                </div>
              </article>
            ))}
          </div>
        )}
      </section>
    );
  }
}